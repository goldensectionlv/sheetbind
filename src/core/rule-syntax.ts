/** Separators inside JSON strings, arrays and objects belong to the argument. */
export function splitRuleText(text: string, separator: string): string[] {
  const parts: string[] = []
  const stack: string[] = []
  let quoted = false
  let escaped = false
  let start = 0
  for (let index = 0; index < text.length; index++) {
    const char = text[index]
    if (quoted) {
      if (escaped) {
        escaped = false
      }
      else if (char === '\\') {
        escaped = true
      }
      else if (char === '"') {
        quoted = false
      }
      continue
    }
    if (char === '"') {
      quoted = true
    }
    else if (char === '[' || char === '{') {
      stack.push(char === '[' ? ']' : '}')
    }
    else if (char === ']' || char === '}') {
      if (stack.pop() !== char) {
        throw new SyntaxError('Unbalanced rule argument')
      }
    }
    else if (char === separator && !stack.length) {
      parts.push(text.slice(start, index).trim())
      start = index + 1
    }
  }
  if (quoted || stack.length) {
    throw new SyntaxError('Unclosed rule argument')
  }
  parts.push(text.slice(start).trim())
  return parts
}

export function parseRuleArgument(text: string): unknown {
  if (!text) {
    throw new SyntaxError('Rule arguments cannot be empty; use "" for an empty string')
  }
  if (/^(?:"|\[|\{|true$|false$|null$|-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$)/.test(text)) {
    return JSON.parse(text)
  }
  if (!/^[\w.$-]+$/.test(text)) {
    throw new SyntaxError('Quote string arguments containing punctuation or spaces')
  }
  return text
}
