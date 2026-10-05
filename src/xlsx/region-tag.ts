import { parseDataReference } from '../core/reference'
import { splitRuleText } from '../core/rule-syntax'

export function parseRegionTag(text: string) {
  const [head, ...tokens] = splitRuleText(text, '|')
  const scope = head.startsWith('with ')
  const source = parseDataReference(scope ? head.slice(5) : head)
  if (tokens.length) {
    throw new SyntaxError('Unknown region option; use the closing marker to set the rectangle')
  }
  return { source, scope, close: scope ? 'with' : head }
}
