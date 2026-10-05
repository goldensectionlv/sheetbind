import { readFile } from 'node:fs/promises'
import ExcelJS from 'exceljs'
import { importWorkbookXlsx } from '../src/index'

export async function openWorkbook(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const book = new ExcelJS.Workbook()
  await book.xlsx.load(Uint8Array.from(bytes).buffer)
  return book
}
export async function saveWorkbook(book: ExcelJS.Workbook): Promise<Buffer> {
  return Buffer.from(await book.xlsx.writeBuffer())
}
export function exampleFile(name: string): Promise<Buffer> {
  return readFile(new URL('../examples/' + name, import.meta.url))
}
export async function editExample(name: string, edit: (book: ExcelJS.Workbook) => void) {
  const book = await openWorkbook(await exampleFile(name))
  edit(book)
  return importWorkbookXlsx(await saveWorkbook(book))
}

export async function importAuthoredWorkbook(author: (book: ExcelJS.Workbook) => void) {
  const book = new ExcelJS.Workbook()
  author(book)
  return importWorkbookXlsx(await saveWorkbook(book))
}
