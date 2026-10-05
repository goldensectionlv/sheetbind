import { defineConfig } from 'vitest/config'
import { fileURLToPath } from 'node:url'

// Workbook model and XLSX boundary checks.
export default defineConfig({
  resolve: { alias: [
    { find: 'sheetbind', replacement: fileURLToPath(new URL('./src/index.ts', import.meta.url)) },
  ] },
  test: {
    environment: 'node',
    include: ['test/**/*.test.ts'],
  },
})
