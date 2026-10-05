import { describe, expect, it } from 'vitest'
import { instantiate, TemplateError } from '../src/core/template'
import type { FieldValue, Fragment, ResolvedFragment, ValueExpression } from '../src/core/template'

type Content = { label: string, settings?: { color: string } }
const value = (id: string, expression: ValueExpression): Fragment<Content> => ({ type: 'value', id, value: expression })
const group = (id: string, children: Fragment<Content>[]): Fragment<Content> => ({ type: 'group', id, content: { label: id }, children })
const repeat = (id: string, path: string, body: Fragment<Content>): Fragment<Content> => ({ type: 'repeat', id, source: { path }, body })
function values(node: ResolvedFragment<Content>): FieldValue[] {
  switch (node.type) {
    case 'value': return [node.value]
    case 'scope': return values(node.body)
    case 'repeat': return node.instances.flatMap(values)
    case 'group': return node.children.flatMap(values)
  }
}

describe('format-independent template composition', () => {
  it('resolves a literal and a binding without any grid or file model', () => {
    expect(values(instantiate(group('g', [value('label', { literal: 'Title' }), value('title', { path: 'title' })]), { title: 'Collection' }))).toEqual(['Title', 'Collection'])
  })

  it('uses the same repeat for single values and composed fragments', () => {
    const data = { items: [{ name: 'Atlas', code: '001' }, { name: 'Beacon', code: '002' }] }
    expect(values(instantiate(repeat('items', 'items', value('name', { path: 'name' })), data))).toEqual(['Atlas', 'Beacon'])
    const card = group('card', [value('name', { path: 'name' }), value('label', { literal: 'Code' }), value('code', { path: 'code' })])
    expect(values(instantiate(repeat('items', 'items', card), data))).toEqual(['Atlas', 'Code', '001', 'Beacon', 'Code', '002'])
  })

  it('composes scopes and nested repeats with explicit root access and concrete origins', () => {
    const template = repeat('items', 'items', group('card', [
      value('title', { path: 'title', from: 'root' }),
      { type: 'scope', id: 'details', source: { path: 'details' }, body: value('code', { path: 'code' }) },
      repeat('notes', 'notes', value('text', { path: 'text' })),
    ]))
    const result = instantiate(template, { title: 'Overview', items: [{ details: { code: '001' }, notes: [{ text: 'Ready' }, { text: 'Next' }] }] })
    expect(values(result)).toEqual(['Overview', '001', 'Ready', 'Next'])
    if (result.type !== 'repeat') {
      throw new Error('Expected repeat')
    }
    const card = result.instances[0]
    if (card.type !== 'group' || card.children[2].type !== 'repeat') {
      throw new Error('Expected nested repeat')
    }
    expect(card.children[2].instances[1].origin).toEqual({
      nodeId: 'text', dataPath: '$data.items[0].notes[1].text',
      iterations: [{ nodeId: 'items', index: 0 }, { nodeId: 'notes', index: 1 }],
    })
  })

  it('does not fall back to a parent/root when a current-context value is missing', () => {
    expect(() => instantiate(repeat('items', 'items', value('name', { path: 'name' })), { name: 'Wrong', items: [{}] })).toThrow(TemplateError)
    expect(values(instantiate(repeat('items', 'items', value('name', { path: 'name', optional: true })), { name: 'Wrong', items: [{}] }))).toEqual([null])
  })

  it.each(['00123', '=SUM(A1:A2) {value}', '_x0041_', '', 0, false, null])('preserves scalar %j', input => {
    expect(values(instantiate(value('v', { path: 'input' }), { input }))).toEqual([input])
  })

  it.each([{}, [], NaN, Infinity, new Date()])('rejects non-scalar binding %#', input => {
    expect(() => instantiate(value('v', { path: 'input' }), { input })).toThrow(TemplateError)
  })

  it('keeps siblings around an empty repeat without synthesizing a record', () => {
    expect(values(instantiate(group('g', [value('head', { literal: 'Header' }), repeat('items', 'items', value('name', { path: 'name' })), value('end', { literal: 'End' })]), { items: [] }))).toEqual(['Header', 'End'])
  })

  it.each([undefined, null, {}, 1])('does not reinterpret an invalid collection as empty %#', items => {
    expect(() => instantiate(repeat('items', 'items', value('name', { path: 'name' })), { items })).toThrow(TemplateError)
  })

  it.each([[null], [1], Array(1)])('rejects malformed and sparse collection items %#', items => {
    expect(() => instantiate(repeat('items', 'items', value('name', { path: 'name' })), { items })).toThrow(TemplateError)
  })

  it('validates empty repeat definitions and identity uniqueness before executing data', () => {
    expect(() => instantiate(repeat('items', 'items', value('bad', { path: 'constructor.name' })), { items: [] })).toThrow(TemplateError)
    expect(() => instantiate(group('g', [value('v', { literal: 1 }), value('v', { literal: 2 })]), {})).toThrow(TemplateError)
  })

  it('guards depth, cycles and expansion before the call stack or memory grows unbounded', () => {
    const cycle: any = { type: 'group', id: 'cycle', content: {}, children: [] }
    cycle.children.push(cycle)
    expect(() => instantiate(cycle, {})).toThrow(TemplateError)
    expect(() => instantiate(group('a', [group('b', [value('c', { literal: 1 })])]), {}, { maxDepth: 2 })).toThrow(TemplateError)
    expect(() => instantiate(repeat('items', 'items', value('v', { path: 'v' })), { items: Array(1_000_000) }, { maxNodes: 10 })).toThrow(TemplateError)
    expect(() => instantiate(value('v', { literal: 1 }), {}, { maxNodes: 0 })).toThrow(RangeError)
  })

  it('supports null-prototype objects but not inherited values or class instances', () => {
    const data = Object.assign(Object.create(null), { value: 'Own' })
    expect(values(instantiate(value('v', { path: 'value' }), data))).toEqual(['Own'])
    expect(() => instantiate(value('v', { path: 'value' }), Object.create({ value: 'Inherited' }))).toThrow(TemplateError)
  })

  it('preserves composition grouping, replay equivalence and independent mutable content', () => {
    const body: Fragment<Content> = { type: 'group', id: 'body', content: { label: 'Card', settings: { color: 'green' } }, children: [value('name', { path: 'name' })] }
    const template = repeat('items', 'items', body)
    const data = { items: [{ name: 'Atlas' }, { name: 'Beacon' }] }
    const snapshot = JSON.stringify({ template, data })
    const result = instantiate(template, data)
    expect(instantiate(template, data)).toEqual(result)
    if (result.type !== 'repeat' || result.instances[0].type !== 'group' || result.instances[1].type !== 'group') {
      throw new Error('Expected groups')
    }
    result.instances[0].content.settings!.color = 'blue'
    expect(result.instances[1].content.settings!.color).toBe('green')
    expect(JSON.stringify({ template, data })).toBe(snapshot)
    expect(values(instantiate(group('outer', [value('a', { literal: 1 }), group('inner', [value('b', { literal: 2 })])]), {}))).toEqual([1, 2])
  })
})
