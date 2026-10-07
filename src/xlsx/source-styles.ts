import type JSZip from 'jszip'
import { setXmlAttributes, setXmlElement, xmlAttributes, xmlElements } from './xml'

/** Keep source style IDs; add only explicit default and text variants used by written cells. */
export async function sourceStyles(source: JSZip) {
  const xml = await source.file('xl/styles.xml')!.async('string')
  const originals = xmlElements(xmlElements(xml, 'cellXfs')[0] ?? '', 'xf')
  const all = [...originals]
  const textStyles = new Map<number, number>()
  let normal: number | undefined
  function textStyle(original = 0): number {
    const cached = textStyles.get(original)
    if (cached !== undefined) {
      return cached
    }
    const id = all.length
    all.push(setXmlAttributes(originals[original] ?? originals[0], { numFmtId: 49, applyNumberFormat: 1 }))
    textStyles.set(original, id)
    return id
  }
  function style(original: number, text = false): number {
    if (text && Number(xmlAttributes(originals[original] ?? originals[0]).numFmtId ?? 0) === 0) {
      return textStyle(original)
    }
    // A nonzero ID prevents a moved General cell from inheriting its new row/column format.
    if (!original && normal === undefined) {
      normal = all.length
      all.push(originals[0])
    }
    return original || normal!
  }
  return { style, textStyle,
    save() {
      if (all.length !== originals.length) {
        source.file('xl/styles.xml', setXmlElement(xml, 'cellXfs', `<cellXfs count="${all.length}">${all.join('')}</cellXfs>`))
      }
    } }
}
