import assert from 'node:assert/strict'
import { execFileSync } from 'node:child_process'
import { copyFile, mkdir, readdir, rm } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const destination = path.resolve(root, 'docs/public/examples')

// Only this generated directory may be cleared; examples/ remains the source.
assert.equal(path.dirname(destination), path.join(root, 'docs', 'public'))
assert.equal(path.basename(destination), 'examples')
await rm(destination, { recursive: true, force: true })

async function copyExamples(relative = ''): Promise<void> {
  const source = path.join(root, 'examples', relative)
  for (const entry of await readdir(source, { withFileTypes: true })) {
    const name = path.join(relative, entry.name)
    if (entry.isDirectory()) {
      await copyExamples(name)
    }
    else if (/\.(xlsx|json)$/.test(entry.name)
      || (['tutorials', 'walkthroughs'].includes(relative) && entry.name.endsWith('.ts'))
      || name === path.join('validation', 'rules.ts')) {
      const target = path.join(destination, name)
      await mkdir(path.dirname(target), { recursive: true })
      await copyFile(path.join(source, entry.name), target)
    }
  }
}

await copyExamples()
execFileSync(process.execPath, ['--import', 'tsx', 'examples/tutorials/run.ts', path.join(destination, 'tutorials')], {
  cwd: root,
  stdio: 'inherit',
})
execFileSync(process.execPath, ['--import', 'tsx', 'examples/walkthroughs/generate.ts', path.join(destination, 'walkthroughs')], {
  cwd: root,
  stdio: 'inherit',
})
console.log('Documentation downloads prepared from examples/')
