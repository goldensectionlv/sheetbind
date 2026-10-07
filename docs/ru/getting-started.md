# Первый отчёт

Создадим XLSX с именем заказчика. Нужны Node.js 22.13 или новее, [pnpm 11.5.2](https://pnpm.io/installation) и Excel для создания или открытия книги. Проверьте версию командой `pnpm --version`: команды ниже рассчитаны на pnpm 11.5.2.

Если нужна установка этой версии pnpm, выполните `npm install --global pnpm@11.5.2`.

## 1. Установите пакет

Эти страницы и скачиваемые файлы описывают ещё не опубликованные изменения. Пакет `0.2.0` из npm им не соответствует. Соберите архив из той же рабочей копии Sheetbind, что и документацию. Из корня репозитория:

```sh
pnpm install --frozen-lockfile
pnpm pack --out ./temp/sheetbind.tgz
```

Создайте папку `sheetbind-example` вне рабочей копии Sheetbind и скопируйте в неё `temp/sheetbind.tgz`. Из этой новой папки установите архив, ExcelJS и средство запуска TypeScript:

```sh
pnpm init --init-type module
pnpm add ./sheetbind.tgz exceljs@4.4.0
pnpm add -D tsx
```

`--init-type module` добавляет `"type": "module"` в `package.json`: примеры используют `import` и `await` на верхнем уровне. Для существующего приложения установите зависимости в нём и используйте его настройки модулей. ExcelJS — peer dependency: его нужно установить рядом с Sheetbind.

Дальнейшие команды выполняйте из `sheetbind-example`. Используйте одну сборку пакета для выдачи и чтения форм.

## 2. Создайте шаблон

Скачайте [first-report-template.xlsx](/examples/tutorials/first-report-template.xlsx) в папку приложения или создайте книгу в Excel:

| Ячейка | Содержимое |
| --- | --- |
| A1 | `Customer` |
| B1 | `{customer.name}` |

Сохраните файл как `first-report-template.xlsx`. Шрифт, границы и ширину колонок можно настроить в Excel.

## 3. Передайте данные и сохраните отчёт

Сохраните этот код как `report.ts` рядом с шаблоном. Можно также [скачать скрипт](/examples/tutorials/first-report.ts) и сохранить его под этим именем.

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('first-report-template.xlsx'))
const data = { customer: { name: 'Sample customer' } }
await writeFile('report.xlsx', await renderWorkbookReport(template, data))
```

Запустите:

```sh
pnpm exec tsx report.ts
```

Откройте `report.xlsx` в той же папке. В A1 остаётся `Customer`, в B1 появляется `Sample customer`.

Если при рендере получили `missing-source` для `$data.customer.name`, проверьте структуру данных. Объект `{ name: 'Sample customer' }` не соответствует тегу `{customer.name}`: в примере выше есть нужный объект `customer`.

Храните размеченный файл как шаблон. После изменения тегов или оформления сохраните его и снова вызовите `importWorkbookXlsx`.

## Следующий шаг

В разделе [«Шаблоны»](./templates.md) добавим таблицу, повтор строк и итог. Если пользователь должен заполнить и вернуть файл, переходите к [заполнению формы](./forms.md).
