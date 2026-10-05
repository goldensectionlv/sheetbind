import { access, readdir, readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { createMarkdownRenderer } from 'vitepress'

interface MarkdownLink {
  line: number
  target: string
}

const HERE = path.dirname(fileURLToPath(import.meta.url))
const ROOT = path.resolve(HERE, '..')
const SKIPPED_DIRECTORIES = new Set(['.git', 'dist', 'node_modules', 'temp'])

function assertInsideRoot(target: string): boolean {
  const relative = path.relative(ROOT, target)
  return relative === '' || (!relative.startsWith('..') && !path.isAbsolute(relative))
}

async function collectMarkdownFiles(directory: string): Promise<string[]> {
  const files: string[] = []
  const entries = await readdir(directory, { withFileTypes: true })

  for (const entry of entries) {
    if (entry.isDirectory() && SKIPPED_DIRECTORIES.has(entry.name)) {
      continue
    }

    const absolute = path.join(directory, entry.name)
    if (entry.isDirectory()) {
      files.push(...await collectMarkdownFiles(absolute))
    }
    else if (entry.isFile() && entry.name.toLowerCase().endsWith('.md') && entry.name !== 'AGENTS.md') {
      files.push(absolute)
    }
  }

  return files
}

async function assertLocaleParity(errors: string[]): Promise<void> {
  const docs = path.join(ROOT, 'docs')
  const russianDocs = path.join(docs, 'ru')
  const english = (await collectMarkdownFiles(docs))
    .map(file => path.relative(docs, file))
    .filter((file) => {
      const [topLevel] = file.split(path.sep)
      return topLevel !== '.vitepress' && topLevel !== 'public' && topLevel !== 'ru'
    })
    .sort()
  const russian = (await collectMarkdownFiles(russianDocs))
    .map(file => path.relative(russianDocs, file))
    .sort()

  if (JSON.stringify(english) !== JSON.stringify(russian)) {
    errors.push(
      'docs and docs/ru must contain the same localized Markdown paths:\n'
      + `en: ${english.join(', ')}\nru: ${russian.join(', ')}`,
    )
  }

  for (const filename of english.filter(name => russian.includes(name))) {
    const [englishMarkdown, russianMarkdown] = await Promise.all([
      readFile(path.join(docs, filename), 'utf8'),
      readFile(path.join(russianDocs, filename), 'utf8'),
    ])
    const englishOutline = headingLevels(englishMarkdown)
    const russianOutline = headingLevels(russianMarkdown)
    if (JSON.stringify(englishOutline) !== JSON.stringify(russianOutline)) {
      errors.push(
        `docs/${filename} and docs/ru/${filename} must have matching heading levels:\n`
        + `en: ${englishOutline.join(', ')}\nru: ${russianOutline.join(', ')}`,
      )
    }
  }
}

function headingLevels(markdown: string): number[] {
  const levels: number[] = []
  let fence: string | undefined

  for (const line of markdown.split(/\r?\n/)) {
    const fenceMatch = line.match(/^\s*(```+|~~~+)/)
    if (fenceMatch) {
      if (!fence) {
        fence = fenceMatch[1][0]
      }
      else if (fence === fenceMatch[1][0]) {
        fence = undefined
      }
      continue
    }
    if (fence) {
      continue
    }

    const heading = line.match(/^(#{1,6})\s+/)
    if (heading) {
      levels.push(heading[1].length)
    }
  }

  return levels
}

function linkDestination(raw: string): string {
  const trimmed = raw.trim()
  if (trimmed.startsWith('<')) {
    const closing = trimmed.indexOf('>')
    return closing >= 0 ? trimmed.slice(1, closing) : trimmed
  }

  return trimmed.split(/\s+(?=["'])/, 1)[0]
}

function markdownLinks(markdown: string): MarkdownLink[] {
  const links: MarkdownLink[] = []
  const lines = markdown.split(/\r?\n/)
  let fence: string | undefined

  for (let index = 0; index < lines.length; index++) {
    const line = lines[index]
    const fenceMatch = line.match(/^\s*(```+|~~~+)/)
    if (fenceMatch) {
      if (!fence) {
        fence = fenceMatch[1][0]
      }
      else if (fence === fenceMatch[1][0]) {
        fence = undefined
      }
      continue
    }
    if (fence) {
      continue
    }

    for (const match of line.matchAll(/!?\[[^\]\n]*\]\(([^)\n]+)\)/g)) {
      links.push({ line: index + 1, target: linkDestination(match[1]) })
    }

    const reference = line.match(/^\s*\[[^\]\n]+\]:\s*(\S+)/)
    if (reference) {
      links.push({ line: index + 1, target: linkDestination(reference[1]) })
    }
  }

  return links
}

function decode(value: string): string {
  try {
    return decodeURIComponent(value)
  }
  catch {
    return value
  }
}

function slugBase(heading: string): string {
  return heading
    .replace(/\s+#+\s*$/, '')
    .replace(/[`*_~]/g, '')
    .toLowerCase()
    .replace(/[^\p{L}\p{M}\p{N}\s_-]/gu, '')
    .replace(/\s/g, '-')
}

async function markdownAnchors(file: string): Promise<Set<string>> {
  const markdown = await readFile(file, 'utf8')
  const docsRoot = path.join(ROOT, 'docs')
  const relative = path.relative(docsRoot, file)
  if (!relative.startsWith('..') && !path.isAbsolute(relative)) {
    const renderer = await createMarkdownRenderer(docsRoot, { highlight: () => '' })
    return new Set(renderer.parse(markdown, {})
      .filter(token => token.type === 'heading_open')
      .map(token => token.attrGet('id'))
      .filter((id): id is string => id !== null))
  }

  const anchors = new Set<string>()
  const occurrences = new Map<string, number>()
  let fence: string | undefined

  for (const line of markdown.split(/\r?\n/)) {
    const fenceMatch = line.match(/^\s*(```+|~~~+)/)
    if (fenceMatch) {
      if (!fence) {
        fence = fenceMatch[1][0]
      }
      else if (fence === fenceMatch[1][0]) {
        fence = undefined
      }
      continue
    }
    if (fence) {
      continue
    }

    const heading = line.match(/^#{1,6}\s+(.+)$/)?.[1]
    if (!heading) {
      continue
    }

    const base = slugBase(heading)
    const count = occurrences.get(base) ?? 0
    occurrences.set(base, count + 1)
    anchors.add(count === 0 ? base : `${base}-${count}`)
  }

  return anchors
}

function isExternal(target: string): boolean {
  return target.startsWith('//') || /^[a-z][a-z\d+.-]*:/i.test(target)
}

async function resolveTarget(source: string, rawPath: string): Promise<string> {
  const decoded = decode(rawPath.split('?', 1)[0])
  if (!decoded.startsWith('/')) {
    return path.resolve(path.dirname(source), decoded)
  }

  const relative = decoded.slice(1)
  const docs = path.join(ROOT, 'docs')
  const candidates = [
    path.join(docs, 'public', relative),
    path.join(docs, `${relative}.md`),
    path.join(docs, relative, 'index.md'),
    path.join(docs, relative),
    path.join(ROOT, relative),
  ]

  for (const candidate of candidates) {
    if (!assertInsideRoot(candidate)) {
      continue
    }

    try {
      await access(candidate)
      return candidate
    }
    catch {
      // Try the next VitePress or repository-local interpretation.
    }
  }

  return candidates[0]
}

async function main(): Promise<void> {
  const files = (await collectMarkdownFiles(ROOT)).sort()
  const anchorCache = new Map<string, Promise<Set<string>>>()
  const errors: string[] = []

  await assertLocaleParity(errors)

  for (const source of files) {
    const markdown = await readFile(source, 'utf8')
    for (const link of markdownLinks(markdown)) {
      if (!link.target || isExternal(link.target)) {
        continue
      }

      const [rawPath, rawFragment] = link.target.split('#', 2)
      const target = rawPath ? await resolveTarget(source, rawPath) : source
      const location = `${path.relative(ROOT, source)}:${link.line}`

      if (!assertInsideRoot(target)) {
        errors.push(`${location} points outside the repository: ${link.target}`)
        continue
      }
      if (source.startsWith(path.join(ROOT, 'docs') + path.sep)
        && !target.startsWith(path.join(ROOT, 'docs') + path.sep)) {
        errors.push(`${location} points outside the published documentation: ${link.target}`)
        continue
      }

      try {
        await access(target)
      }
      catch {
        errors.push(`${location} has a missing target: ${link.target}`)
        continue
      }

      if (!rawFragment || path.extname(target).toLowerCase() !== '.md') {
        continue
      }

      let anchors = anchorCache.get(target)
      if (!anchors) {
        anchors = markdownAnchors(target)
        anchorCache.set(target, anchors)
      }
      const fragment = decode(rawFragment).toLowerCase()
      if (!(await anchors).has(fragment)) {
        errors.push(`${location} has a missing heading: ${link.target}`)
      }
    }
  }

  if (errors.length > 0) {
    throw new Error(`documentation links failed:\n${errors.join('\n')}`)
  }

  console.log(`documentation: ${files.length} Markdown files checked (links, anchors, locale parity)`)
}

await main()
