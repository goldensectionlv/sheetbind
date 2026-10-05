# Sheetbind

[English](./README.md) | [Русский](./README.ru.md)

Sheetbind заполняет XLSX-шаблоны данными. Разметьте ячейки тегами в Excel, передайте файл и данные в код и сохраните отчёт. Для ввода данных создайте форму: пользователь заполнит её в Excel, а приложение прочитает результат.

Нужны Node.js 22.13 или новее и ExcelJS 4.4.0. Лицензия MIT.

## Первый отчёт

Создайте `first-report-template.xlsx`: в A1 запишите `Customer`, в B1 — `{customer.name}`. Шрифты, цвета и ширину колонок настройте в Excel.

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('first-report-template.xlsx'))
const data = { customer: { name: 'Sample customer' } }
await writeFile('report.xlsx', await renderWorkbookReport(template, data))
```

В результате B1 содержит `Sample customer`. В инструкции [«Первый отчёт»](./docs/ru/getting-started.md) есть установка пакета, готовый шаблон и команда запуска.

## Руководство

- [Шаблоны](./docs/ru/templates.md): значения, повтор строк, вложенные данные и повтор колонок.
- [Проверки и списки](./docs/ru/fields.md): правила полей и выпадающие списки.
- [Заполнение формы](./docs/ru/forms.md): выдача XLSX, редактирование и чтение данных.
- [API](./docs/ru/api.md): функции, параметры, результаты и ошибки.
- [Excel и ограничения](./docs/ru/xlsx.md): оформление, формулы и поддерживаемые возможности книги.

Храните размеченный XLSX как шаблон. После редактирования загрузите файл заново через `importWorkbookXlsx`. Формулы пересчитывает Excel; Sheetbind не вычисляет их результат.

Для работы над самой библиотекой есть [разработка](./docs/ru/development.md) и [архитектура](./docs/ru/architecture.md).
