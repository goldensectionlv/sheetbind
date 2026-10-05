/** Small edits to known OOXML elements; all other markup stays in its source part. */
export function xmlAttributes(tag: string): Record<string, string> {
  return Object.fromEntries([...tag.matchAll(/([\w:.-]+)\s*=\s*(["'])(.*?)\2/g)].map(match => [match[1], decodeXml(match[3])]))
}
export function decodeXml(value: string): string {
  return value.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-f]+);/gi, entity => ({ '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" })[entity]
    ?? String.fromCodePoint(entity[2].toLowerCase() === 'x' ? parseInt(entity.slice(3, -1), 16) : Number(entity.slice(2, -1))))
}
export function encodeXml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}
export function xmlElements(xml: string, name: string): string[] {
  return xml.match(new RegExp(`<${name}\\b[^>]*?(?:/>|>[\\s\\S]*?</${name}>)`, 'g')) ?? []
}
export function xmlBody(xml: string): string {
  return xml.replace(/^<[^>]*>/, '').replace(/<\/[^>]*>$/, '')
}
export function setXmlAttributes(xml: string, values: Readonly<Record<string, string | number | undefined>>): string {
  return xml.replace(/^<[^>]+>/, head => {
    for (const [key, value] of Object.entries(values)) {
      const pattern = new RegExp(`\\s${key}\\s*=\\s*(["']).*?\\1`)
      head = head.replace(pattern, '')
      if (value !== undefined) {
        head = head.replace(/\/?>(?=$)/, end => ` ${key}="${encodeXml(String(value))}"${end}`)
      }
    }
    return head
  })
}
export function setXmlElement(xml: string, name: string, content: string): string {
  const found = xmlElements(xml, name)[0]
  if (found) {
    return xml.replace(found, () => content)
  }
  return content ? xml.replace(/<\/[^>]+>\s*$/, end => content + end) : xml
}
export function appendXmlChildren(xml: string, name: string, children: readonly string[], count = false): string {
  if (!children.length) {
    return xml
  }
  const found = xmlElements(xml, name)[0]
  const body = found && !found.endsWith('/>') ? xmlBody(found) : ''
  const total = Number(found ? xmlAttributes(found.split('>')[0]).count ?? 0 : 0) + children.length
  let head = found ? found.split('>')[0].replace(/\/$/, '') + '>' : `<${name}>`
  if (count) {
    head = setXmlAttributes(head, { count: total })
  }
  return setXmlElement(xml, name, `${head}${body}${children.join('')}</${name}>`)
}
export function resolvePart(part: string, target: string): string {
  const parts = target.startsWith('/') ? [] : part.split('/').slice(0, -1)
  for (const segment of target.split('/')) {
    if (segment === '..') {
      parts.pop()
    }
    else if (segment && segment !== '.') {
      parts.push(segment)
    }
  }
  return parts.join('/')
}
