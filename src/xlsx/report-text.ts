import type JSZip from 'jszip'

/** Decode ST_Xstring once: protected literal escapes must not be decoded recursively. */
export function decodeXstring(value: string): string {
  return value.replace(/_x([0-9a-fA-F]{4})_/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)))
}

export function hasInvalidXmlText(value: string): boolean {
  // Unicode mode matches lone surrogates without rejecting valid astral characters.
  // eslint-disable-next-line no-control-regex
  return /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff\ufffe\uffff]/u.test(value)
}

/** The same serialized text boundary applies to configured and tagged reports. */
export function xlsxTextIssues(value: string): { code: string, message: string }[] {
  const issues: { code: string, message: string }[] = []
  if (value.length > 32767) {
    issues.push({ code: 'text-limit', message: 'cell text exceeds 32767 characters' })
  }
  if (hasInvalidXmlText(value)) {
    issues.push({ code: 'invalid-text', message: 'text contains characters that XML cannot represent' })
  }
  return issues
}

/** Apply the same text limits to every XLSX writer, including hidden lookup cells. */
export function assertXlsxText(value: string, location?: string): void {
  const issue = xlsxTextIssues(value)[0]
  if (issue) {
    throw new RangeError(`${location ? location + ': ' : ''}${issue.message}`)
  }
}

/**
 * ExcelJS 4.4 writes literal `_xHHHH_` sequences without protecting the initial
 * underscore. Excel decodes them in ST_Xstring fields, changing user text.
 * Apply once to a freshly serialized report, never to arbitrary imported XML
 * (where the same sequences may be intentional encoded characters).
 */
export async function escapeReportText(zip: JSZip): Promise<void> {
  for (const [path, part] of Object.entries(zip.files)) {
    const text = path === 'xl/sharedStrings.xml' || /^xl\/worksheets\/sheet\d+\.xml$/.test(path)
    const attribute = path === 'xl/workbook.xml'
      ? 'sheet'
      : path === 'xl/styles.xml' ? 'numFmt' : undefined
    if (!text && !attribute) {
      continue
    }
    const original = await part.async('string')
    let xml = original
    if (text) {
      xml = xml.replace(/(<t(?:\s[^>]*)?>)([\s\S]*?)(<\/t>)/g, (_, start, value, end) => start + protect(value) + end)
    }
    if (path === 'xl/workbook.xml') {
      // Excel also decodes escapes in defined-name formulas. Without protection,
      // an anchor to Sheet_x0041_ becomes an external reference to SheetA.
      xml = xml.replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, (_, start, value, end) => start + protect(value) + end)
    }
    if (attribute) {
      const name = attribute === 'numFmt' ? 'formatCode' : 'name'
      const pattern = new RegExp(`(<${attribute}\\b[^>]*\\b${name}=")([^"]*)(")`, 'g')
      xml = xml.replace(pattern, (_, start, value, end) => start + protect(value) + end)
    }
    if (xml !== original) {
      zip.file(path, xml)
    }
  }
}

export function protect(value: string): string {
  // Lookahead handles sequences sharing an underscore. Uppercase F also avoids
  // ExcelJS's case-sensitive escape reader mishandling the protection itself.
  return value.replace(/_(?=x[0-9a-fA-F]{4}_)/g, '_x005F_')
}
