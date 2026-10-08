# Sheetbind

[English](./README.md) | [Русский](./README.ru.md)

Sheetbind заполняет XLSX-шаблоны данными. Разметьте ячейки тегами в Excel, передайте файл и данные в код и сохраните отчёт. Для ввода данных создайте форму: пользователь заполнит её в Excel, а приложение прочитает результат.

Нужен Node.js 22.13 или новее. Лицензия MIT.

[Документация](https://goldensectionlv.github.io/sheetbind/ru/) содержит примеры с изображениями и книги для скачивания.

Установите пакет в своём приложении:

```sh
npm install sheetbind
```

ExcelJS устанавливается автоматически как зависимость. Sheetbind включает типы TypeScript и поддерживает `import` и `require`; используйте существующие настройки модулей и сборки приложения. Для CommonJS: `const { importWorkbookXlsx, renderWorkbookReport } = require('sheetbind')`.

В инструкции [«Первый отчёт»](https://goldensectionlv.github.io/sheetbind/ru/getting-started) есть готовый шаблон и полный пример. Документация и скачиваемые примеры используют Sheetbind 0.3.1.

## Первый отчёт

Создайте `first-report-template.xlsx`: в A1 запишите `Customer`, в B1 — `{customer.name}`. Шрифты, цвета и ширину колонок настройте в Excel.

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

async function main() {
  const template = await importWorkbookXlsx(await readFile('first-report-template.xlsx'))
  const data = { customer: { name: 'Sample customer' } }
  await writeFile('report.xlsx', await renderWorkbookReport(template, data))
}

main().catch(error => {
  console.error(error)
  process.exitCode = 1
})
```

В результате B1 содержит `Sample customer`. В инструкции [«Первый отчёт»](https://goldensectionlv.github.io/sheetbind/ru/getting-started) есть установка пакета, готовый шаблон и команда запуска.

## Руководство

- [Шаблоны](https://goldensectionlv.github.io/sheetbind/ru/templates): значения, повтор строк, вложенные данные и повтор колонок.
- [Проверки и списки](https://goldensectionlv.github.io/sheetbind/ru/fields): правила полей и выпадающие списки.
- [Заполнение формы](https://goldensectionlv.github.io/sheetbind/ru/forms): выдача XLSX, редактирование и чтение данных.
- [API](https://goldensectionlv.github.io/sheetbind/ru/api): функции, параметры, результаты и ошибки.
- [Excel и ограничения](https://goldensectionlv.github.io/sheetbind/ru/xlsx): оформление, формулы и поддерживаемые возможности книги.

Храните размеченный XLSX как шаблон. После редактирования загрузите файл заново через `importWorkbookXlsx`. Формулы пересчитывает Excel; Sheetbind не вычисляет их результат.

Для работы над самой библиотекой есть [разработка](https://goldensectionlv.github.io/sheetbind/ru/development) и [архитектура](https://goldensectionlv.github.io/sheetbind/ru/architecture).
