import { defineConfig } from 'tsup'

// Dual ESM + CJS build with type declarations. esbuild (tsup) bundles the
// extensionless relative imports, so the source needs no `.js` suffixes. ExcelJS
// remains the consumer's peer dependency in both the runtime and declarations.
export default defineConfig({
  clean: true,
  dts: true,
  entry: ['src/index.ts'],
  // Let Rollup emit CJS chunks and maps; tsup's CJS transform leaks absolute paths.
  treeshake: true,
  external: ['exceljs'],
  format: ['esm', 'cjs'],
  sourcemap: true,
  target: 'node22',
})
