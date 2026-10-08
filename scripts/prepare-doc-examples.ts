import assert from 'node:assert/strict'
import { execFileSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { copyFile, mkdir, readdir, readFile, rm } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const docs = path.join(root, 'docs')
const destination = path.resolve(root, 'docs/public/examples')
const generated = path.resolve(root, 'temp/documentation-examples')
const downloads = new Set<string>()

assert.equal(path.dirname(destination), path.join(root, 'docs', 'public'))
assert.equal(path.basename(destination), 'examples')
await rm(destination, { recursive: true, force: true })
assert.equal(path.dirname(generated), path.join(root, 'temp'))
assert.equal(path.basename(generated), 'documentation-examples')
await rm(generated, { recursive: true, force: true })

async function collectDownloads(directory: string): Promise<void> {
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const source = path.join(directory, entry.name)
    if (entry.isDirectory() && entry.name !== 'public' && !entry.name.startsWith('.')) {
      await collectDownloads(source)
    }
    else if (entry.isFile() && entry.name.endsWith('.md')) {
      const markdown = await readFile(source, 'utf8')
      for (const [, name] of markdown.matchAll(/\/examples\/([\w./-]+\.(?:xlsx|json|ts))/g)) {
        assert(name.split('/').every(part => part && part !== '.' && part !== '..'), `Invalid example path: ${name}`)
        downloads.add(name)
      }
    }
  }
}

await collectDownloads(docs)
assert(downloads.size > 0, 'Documentation must reference its downloadable examples')
execFileSync(process.execPath, ['--import', 'tsx', 'examples/tutorials/run.ts', path.join(generated, 'tutorials')], {
  cwd: root,
  stdio: 'inherit',
})
execFileSync(process.execPath, ['--import', 'tsx', 'examples/walkthroughs/generate.ts', path.join(generated, 'walkthroughs')], {
  cwd: root,
  stdio: 'inherit',
})
for (const name of [...downloads].sort()) {
  const built = path.join(generated, name)
  const source = existsSync(built) ? built : path.join(root, 'examples', name)
  const target = path.join(destination, name)
  await mkdir(path.dirname(target), { recursive: true })
  await copyFile(source, target)
}
console.log(`Documentation downloads: ${downloads.size} files referenced by the guide`)
