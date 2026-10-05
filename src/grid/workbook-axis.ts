import type { ResolvedFragment } from '../core/template'
import { WorkbookAxis, workbookIssue, WORKBOOK_LIMITS } from './workbook'
import type { WorkbookBody } from './workbook'
import { FormulaEdge } from './workbook-formula'

export interface AxisPlan { readonly size: number, readonly segments: readonly AxisSegment[] }
interface AxisSegment {
  readonly start: number
  readonly length: number
  readonly ids: readonly string[]
  readonly instances: readonly AxisPlan[]
  readonly offsets: readonly number[]
}
interface AxisDraft {
  readonly length: number
  readonly segments: { start: number, length: number, ids: string[], path: string, instances: AxisDraft[] }[]
}
export type WorkbookAxes = Readonly<Record<WorkbookAxis, AxisPlan>>
export const workbookAxes = {
  [WorkbookAxis.Rows]: { position: 'row', length: 'height', coordinate: 'row' },
  [WorkbookAxis.Columns]: { position: 'column', length: 'width', coordinate: 'column' },
} as const

function emptyAxis(length = 0): AxisDraft {
  return { length, segments: [] }
}

/** Fixed content retains its grid line even inside a neighbouring nested repeat. */
function retainFixedSpan(plan: AxisDraft, start: number, end: number): void {
  for (const segment of plan.segments) {
    if (start >= segment.start + segment.length || segment.start >= end) {
      continue
    }
    const first = segment.instances[0] ??= emptyAxis(segment.length)
    retainFixedSpan(first, Math.max(1, start - segment.start + 1), Math.min(segment.length + 1, end - segment.start + 1))
  }
}

/** Each axis shares geometric bands across the other axis. Data remains in core. */
export function planWorkbookAxes(body: WorkbookBody, resolved: ResolvedFragment<null>, key: (id: string) => string): WorkbookAxes {
  function planAxis(axis: WorkbookAxis): AxisPlan {
    const root = emptyAxis()
    const shape = workbookAxes[axis]
    const fixed: { plan: AxisDraft, definition: WorkbookBody, offset: number }[] = []
    function collect(plan: AxisDraft, definition: WorkbookBody, node: ResolvedFragment<null>, offset: number): void {
      if (node.type !== 'group') {
        throw new Error('Expected a workbook body')
      }
      const nodes = new Map(node.children.map(child => [child.origin.nodeId, child]))
      for (const region of definition.regions ?? []) {
        const child = nodes.get(key(region.id))!
        const instances = child.type === 'repeat' ? child.instances : child.type === 'scope' ? [child.body] : []
        const start = offset + (region[shape.position] ?? 1)
        const length = region[shape.length] ?? WORKBOOK_LIMITS[axis]
        if (region.type === 'repeat' && (region.axis ?? WorkbookAxis.Rows) === axis) {
          let segment = plan.segments.find(value => value.start === start && value.length === length)
          if (!segment) {
            if (plan.segments.some(value => start < value.start + value.length && value.start < start + length)) {
              workbookIssue('growth-band-overlap', region.id, `Overlapping ${axis} repeats must reserve the same band`, region.id)
            }
            segment = { start, length, ids: [], path: child.origin.dataPath, instances: [] }
            plan.segments.push(segment)
          }
          if (!segment.ids.includes(region.id)) {
            segment.ids.push(region.id)
          }
          for (const [index, instance] of instances.entries()) {
            collect(segment.instances[index] ??= emptyAxis(length), region, instance, 0)
          }
        }
        else {
          for (const instance of instances) {
            collect(plan, region, instance, start - 1)
          }
        }
      }
      fixed.push({ plan, definition, offset })
    }
    collect(root, body, resolved, 0)
    for (const { plan, definition, offset } of fixed) {
      if (!plan.segments.length) {
        continue
      }
      for (const cell of definition.cells) {
        const start = offset + cell.at[shape.coordinate]
        retainFixedSpan(plan, start, start + cell.size[axis])
      }
      for (const range of definition.occupied ?? []) {
        retainFixedSpan(plan, offset + range.start[shape.coordinate], offset + range.end[shape.coordinate] + 1)
      }
    }
    return completeAxis(root, axis)
  }
  return { rows: planAxis(WorkbookAxis.Rows), columns: planAxis(WorkbookAxis.Columns) }
}

/** Readers receive complete geometry; offsets are computed once for every consumer. */
function completeAxis(draft: AxisDraft, axis: WorkbookAxis): AxisPlan {
  const segments = [...draft.segments].sort((a, b) => a.start - b.start).map(segment => {
    const instances = segment.instances.map(instance => completeAxis(instance, axis))
    const offsets = [0]
    for (const instance of instances) {
      offsets.push(offsets.at(-1)! + instance.size)
    }
    if (offsets.at(-1)! > WORKBOOK_LIMITS[axis]) {
      workbookIssue(`${axis === WorkbookAxis.Rows ? 'row' : 'column'}-limit`, segment.ids[0], `Repeated ${axis} exceed the XLSX limit`, segment.path, 'data')
    }
    return { start: segment.start, length: segment.length, ids: segment.ids, instances, offsets }
  })
  return { segments, size: draft.length + segments.reduce((sum, segment) => sum + segment.offsets.at(-1)! - segment.length, 0) }
}

function iteration(segment: AxisSegment, context: ReadonlyMap<string, number>, fallback: number): number {
  for (const id of segment.ids) {
    if (context.has(id)) {
      return context.get(id)!
    }
  }
  return fallback
}

export function mapAxis(plan: AxisPlan, position: number, edge = FormulaEdge.Cell, context: ReadonlyMap<string, number> = new Map()): number | undefined {
  let growth = 0
  for (const segment of plan.segments) {
    if (position < segment.start) {
      break
    }
    const offsets = segment.offsets
    if (position >= segment.start + segment.length) {
      growth += offsets.at(-1)! - segment.length
      continue
    }
    const start = segment.start + growth
    if (!segment.instances.length) {
      return edge === FormulaEdge.Cell ? undefined : start - Number(edge === FormulaEdge.End)
    }
    const index = iteration(segment, context, edge === FormulaEdge.End ? segment.instances.length - 1 : 0)
    const child = segment.instances[index]
    if (!child) {
      throw new RangeError(`Missing placement for ${segment.ids[0]}[${index}]`)
    }
    const local = mapAxis(child, position - segment.start + 1, edge, context)
    return local === undefined ? undefined : start - 1 + offsets[index] + local
  }
  return position + growth
}

/** Metadata at one authored coordinate applies to every copy of that grid line. */
export function axisPositions(plan: AxisPlan, position: number): number[] {
  let growth = 0
  for (const segment of plan.segments) {
    if (position < segment.start) {
      break
    }
    if (position >= segment.start + segment.length) {
      growth += segment.offsets.at(-1)! - segment.length
      continue
    }
    return segment.instances.flatMap((instance, index) => axisPositions(instance, position - segment.start + 1)
      .map(value => value + segment.start + growth - 1 + segment.offsets[index]))
  }
  return [position + growth]
}

/** Copy-relative references follow common bands; unrelated references keep the copy delta. */
export function sharedAxisContext(plan: AxisPlan, source: number, target: number, indexes: ReadonlyMap<string, number>): Map<string, number> {
  const segment = plan.segments.find(item => source >= item.start && source < item.start + item.length && target >= item.start && target < item.start + item.length)
  if (!segment?.instances.length) {
    return new Map()
  }
  const index = iteration(segment, indexes, 0)
  const context = sharedAxisContext(segment.instances[index], source - segment.start + 1, target - segment.start + 1, indexes)
  for (const id of segment.ids) {
    context.set(id, index)
  }
  return context
}
