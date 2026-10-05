import type { ChoiceOption, ResolvedChoice } from '../core/choices'

export interface WorkbookChoiceOption extends ChoiceOption { readonly text: string }
export interface WorkbookChoice { readonly text: string | null, readonly items: readonly WorkbookChoiceOption[] }

/** The canvas and XLSX show the same unambiguous text; core keeps the original labels. */
export function createWorkbookChoiceDisplay() {
  const sources = new WeakMap<readonly ChoiceOption[], readonly WorkbookChoiceOption[]>()
  return (choice: ResolvedChoice): WorkbookChoice => {
    let items = sources.get(choice.items)
    if (!items) {
      const counts = new Map<string, number>()
      const texts = new Set<string>()
      choice.items.forEach(option => counts.set(option.label.toLowerCase(), (counts.get(option.label.toLowerCase()) ?? 0) + 1))
      items = choice.items.map(option => {
        const text = counts.get(option.label.toLowerCase())! > 1 ? `${option.label} [${option.key}]` : option.label
        if (texts.has(text.toLowerCase())) {
          throw new SyntaxError('Choice display labels are ambiguous after adding keys')
        }
        texts.add(text.toLowerCase())
        return { ...option, text }
      })
      sources.set(choice.items, items)
    }
    return { items, text: choice.key === null ? null : items.find(item => item.key === choice.key)?.text ?? String(choice.key) }
  }
}
