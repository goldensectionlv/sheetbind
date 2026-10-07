import { defineConfig } from 'tsup'
import { writeFile } from 'node:fs/promises'
import { createRequire } from 'node:module'

// Both entry points share one runtime and one declaration of the opaque template.
export default defineConfig({
  clean: true,
  dts: true,
  entry: ['src/index.ts'],
  // Let Rollup emit CJS chunks and maps; tsup's CJS transform leaks absolute paths.
  treeshake: true,
  external: ['exceljs'],
  format: ['cjs'],
  sourcemap: true,
  target: 'node22',
  async onSuccess() {
    const require = createRequire(import.meta.url)
    const entry = require.resolve('./dist/index.cjs')
    delete require.cache[entry]
    const names = Object.keys(require(entry)).sort()
    await writeFile('dist/index.js', `import api from './index.cjs'\nexport const { ${names.join(', ')} } = api\n`)
  },
})
