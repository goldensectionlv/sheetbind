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

export function protect(value: string): string {
  // Lookahead handles sequences sharing an underscore. Uppercase F also avoids
  // ExcelJS's case-sensitive escape reader mishandling the protection itself.
  return value.replace(/_(?=x[0-9a-fA-F]{4}_)/g, '_x005F_')
}
