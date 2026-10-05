import type JSZip from 'jszip'
import { appendXmlChildren, setXmlAttributes, setXmlElement, xmlAttributes, xmlElements } from './xml'

/** Append generated styles to the source tables; original style IDs and extensions remain valid. */
export async function sourceStyles(source: JSZip, generated: JSZip) {
  let xml = await source.file('xl/styles.xml')!.async('string')
  const fresh = await generated.file('xl/styles.xml')!.async('string')
  const entries = (document: string, parent: string, name: string) => xmlElements(xmlElements(document, parent)[0] ?? '', name)
  const offsets = new Map<string, number>()
  for (const [parent, name] of [['fonts', 'font'], ['fills', 'fill'], ['borders', 'border']] as const) {
    offsets.set(name, entries(xml, parent, name).length)
    xml = appendXmlChildren(xml, parent, entries(fresh, parent, name), true)
  }
  const numFmts = new Map<number, number>()
  const formats = entries(xml, 'numFmts', 'numFmt')
  let nextFormat = Math.max(163, ...formats.map(value => Number(xmlAttributes(value).numFmtId))) + 1
  xml = appendXmlChildren(xml, 'numFmts', entries(fresh, 'numFmts', 'numFmt').map(value => {
    const id = nextFormat++
    numFmts.set(Number(xmlAttributes(value).numFmtId), id)
    return setXmlAttributes(value, { numFmtId: id })
  }), true)
  const originals = entries(xml, 'cellXfs', 'xf')
  const added = entries(fresh, 'cellXfs', 'xf').map(value => {
    const attr = xmlAttributes(value.split('>')[0])
    return setXmlAttributes(value, { fontId: Number(attr.fontId ?? 0) + offsets.get('font')!, fillId: Number(attr.fillId ?? 0) + offsets.get('fill')!,
      borderId: Number(attr.borderId ?? 0) + offsets.get('border')!, numFmtId: numFmts.get(Number(attr.numFmtId)) ?? Number(attr.numFmtId ?? 0), xfId: 0 })
  })
  const all = [...originals, ...added]
  const textStyles = new Map<number, number>()
  const normal = all.length
  all.push(originals[0])
  function style(original: number, text = false): number {
    const value = originals[original] ?? originals[0]
    if (!text || Number(xmlAttributes(value).numFmtId ?? 0) !== 0) {
      return original || normal
    }
    const cached = textStyles.get(original)
    if (cached !== undefined) {
      return cached
    }
    const id = all.length
    all.push(setXmlAttributes(value, { numFmtId: 49, applyNumberFormat: 1 }))
    textStyles.set(original, id)
    return id
  }
  return { generated: (index: number) => originals.length + index, style,
    save() {
      xml = setXmlElement(xml, 'cellXfs', `<cellXfs count="${all.length}">${all.join('')}</cellXfs>`)
      const formats = xmlElements(xml, 'numFmts')[0]
      if (formats) {
        xml = xml.replace(formats, '').replace(/<fonts\b/, () => formats + '<fonts')
      }
      source.file('xl/styles.xml', xml)
    } }
}
