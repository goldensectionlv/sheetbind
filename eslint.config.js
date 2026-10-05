import js from '@eslint/js'
import stylistic from '@stylistic/eslint-plugin'
import { defineConfig } from 'eslint/config'
import tseslint from 'typescript-eslint'

export default defineConfig(
  { ignores: ['coverage/**', 'dist/**', 'node_modules/**', 'temp/**'] },
  {
    files: ['src/**/*.ts', 'test/**/*.ts', 'scripts/**/*.ts', 'examples/**/*.ts', '*.config.ts', 'eslint.config.js'],
    extends: [js.configs.recommended, tseslint.configs.recommended],
    plugins: { '@stylistic': stylistic },
    rules: {
      '@typescript-eslint/no-explicit-any': 'off',
      'curly': ['error', 'all'],
      'one-var': ['error', 'never'],
      'no-sequences': 'error',
      '@stylistic/brace-style': ['error', 'stroustrup', { allowSingleLine: false }],
      '@stylistic/max-statements-per-line': ['error', { max: 1 }],
      '@stylistic/indent': ['error', 2, { SwitchCase: 1 }],
      '@stylistic/semi': ['error', 'never'],
      '@stylistic/quotes': ['error', 'single', { avoidEscape: true }],
      '@stylistic/comma-dangle': ['error', 'always-multiline'],
      '@stylistic/comma-spacing': 'error',
      '@stylistic/comma-style': ['error', 'last'],
      '@stylistic/array-bracket-spacing': ['error', 'never'],
      '@stylistic/object-curly-spacing': ['error', 'always'],
      '@stylistic/key-spacing': 'error',
      '@stylistic/keyword-spacing': 'error',
      '@stylistic/arrow-spacing': 'error',
      '@stylistic/function-call-spacing': ['error', 'never'],
      '@stylistic/space-before-function-paren': ['error', { anonymous: 'always', named: 'never', asyncArrow: 'always' }],
      '@stylistic/space-before-blocks': 'error',
      '@stylistic/space-in-parens': ['error', 'never'],
      '@stylistic/space-infix-ops': 'error',
      '@stylistic/type-annotation-spacing': 'error',
      '@stylistic/member-delimiter-style': ['error', {
        multiline: { delimiter: 'none', requireLast: false },
        singleline: { delimiter: 'comma', requireLast: false },
      }],
      '@stylistic/multiline-ternary': ['error', 'always-multiline'],
      '@stylistic/no-multi-spaces': 'error',
      '@stylistic/no-trailing-spaces': 'error',
      '@stylistic/no-multiple-empty-lines': ['error', { max: 1, maxBOF: 0, maxEOF: 0 }],
      '@stylistic/eol-last': ['error', 'always'],
    },
  },
  {
    files: ['src/**/*.ts'],
    rules: { '@typescript-eslint/no-explicit-any': 'error' },
  },
  {
    files: ['src/grid/geometry.ts'],
    rules: { 'no-restricted-imports': ['error', { patterns: [{ group: ['*'], message: 'Geometry remains dependency-free.' }] }] },
  },
  {
    files: ['src/core/**/*.ts'],
    rules: { 'no-restricted-imports': ['error', { patterns: [{ regex: '^(?!\\./(?:template|json|reference|field-rules|validation|rule-syntax|dictionaries|choices)$)', message: 'Core semantics may depend only on core primitives.' }] }] },
  },
  {
    files: ['src/grid/workbook*.ts'],
    rules: { 'no-restricted-imports': ['error', { patterns: [{ regex: '^(?!\\./(?:geometry|workbook-choice-display|workbook(?:-[a-z]+)*)$|\\.\\./core/(?:template|json|reference|field-rules|validation|rule-syntax|dictionaries|choices)$)', message: 'Workbook layout depends on core and grid, not XLSX or UI.' }] }] },
  },
  {
    files: ['src/form/**/*.ts'],
    rules: { 'no-restricted-imports': ['error', { patterns: [{ regex: '^(?!\\./(?:records|workbook(?:-[a-z]+)*)$|\\.\\./(?:core|grid)/[a-z-]+$)', message: 'Form semantics must not depend on XLSX transport or UI.' }] }] },
  },
  {
    files: ['src/xlsx/**/*.ts'],
    rules: { 'no-restricted-imports': ['error', { patterns: [{ regex: '^(?!\\./[a-z-]+$|\\.\\./(?:core|grid|form)/[a-z-]+$|(?:exceljs|jszip|node:crypto)$)', message: 'XLSX transport uses core/grid/form and explicit I/O dependencies.' }] }] },
  },
)
