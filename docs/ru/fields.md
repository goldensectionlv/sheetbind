# Проверки и списки

Правила записываются в той же ячейке, что и [привязка](./templates.md):

```text
{quantity}{@validate:required|number|min:0}
```

Здесь `quantity` — поле данных, `required` запрещает пустое значение, `number` требует число, `min:0` запрещает отрицательные числа. Правила выполняются при построении отчёта и чтении заполненной формы. При выдаче формы проверка заполнения откладывается до чтения.

## Встроенные правила

В библиотеке восемь встроенных правил:

| Правило | Что проверяет | Пример |
| --- | --- | --- |
| `required` | Значение заполнено | `required` |
| `string` | Строка | `string` |
| `number` | Конечное число | `number` |
| `boolean` | `true` или `false` | `boolean` |
| `object` | Объект, не массив | `object` |
| `min` | Число не меньше границы | `min:0` |
| `max` | Число не больше границы | `max:24` |
| `maxLength` | Длина строки не больше границы | `maxLength:100` |

`min` и `max` принимают одно конечное число. Для `maxLength` нужна неотрицательная целая граница; длина считается в кодовых единицах UTF-16.

Правила идут слева направо. Первая ошибка останавливает проверку этого поля; остальные поля продолжают проверяться. Несколько директив `@validate` в одной ячейке добавляют правила в указанном порядке.

Пустыми считаются `null`, `""` и строки из пробелов. Все встроенные правила, кроме `required`, пропускают пустые значения. `0` и `false` заполнены. Проверки не преобразуют значения: строка `"12"` не проходит `number`, а пробелы не обрезаются.

Если поле необязательно, не добавляйте `required`. Знак `?` в привязке разрешает отсутствие поля во входных данных отчёта; он не отменяет `required` при чтении формы. Если значения нет, пропустите свойство объекта: явный `undefined` недопустим во входных JSON-данных.

## Список строк

Когда нужно выбрать сам текст, используйте `@list`:

```text
{status}{@list:Statuses}
```

Создайте лист `Input` в `list-template.xlsx`: A1 = `Status`, B1 = этот тег. Или скачайте [готовый шаблон](/examples/tutorials/list-template.xlsx). После [установки](./getting-started.md) сохраните рядом [list-report.ts](/examples/tutorials/list-report.ts):

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('list-template.xlsx'))
const data = { status: 'Draft' }
const dictionaries = { Statuses: ['Draft', 'Ready'] }
await writeFile('list-report.xlsx', await renderWorkbookReport(template, data, { dictionaries }))
```

Из этой папки выполните:

```sh
pnpm exec tsx list-report.ts
```

В `list-report.xlsx` ячейка B1 содержит `Draft` и предлагает список `Draft`, `Ready`. Чтобы обязать пользователя выбрать значение, добавьте `{@validate:required}`.

Этот пример создаёт отчёт. Для файла, который пользователь вернёт приложению, замените `renderWorkbookReport` на `renderWorkbookForm` в импорте и вызове. При чтении такой [формы](./forms.md) поле возвращает выбранную строку. Сам по себе выпадающий список не делает отчёт пригодным для `readWorkbookForm`.

Список должен содержать хотя бы одну строку. При чтении формы передайте тот же список через `dictionaries`. Значение вне списка даст ошибку `list`.

## Выбор объекта или ключа

`@choice` нужен, когда у варианта есть отдельные идентификатор и подпись:

```text
{product}{@choice:Products; key=id; label=name}
```

Создайте лист `Input` в `choice-template.xlsx`: A1 = `Product`, B1 = этот тег. Или скачайте [готовый шаблон](/examples/tutorials/choice-template.xlsx). Сохраните рядом [choice-report.ts](/examples/tutorials/choice-report.ts):

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('choice-template.xlsx'))
const dictionaries = {
  Products: [
    { id: '001', name: 'Paper' },
    { id: '002', name: 'Pen' },
  ],
}
const data = { product: dictionaries.Products[0] }
await writeFile('choice-report.xlsx', await renderWorkbookReport(template, data, { dictionaries }))
```

```sh
pnpm exec tsx choice-report.ts
```

В `choice-report.xlsx` ячейка B1 содержит `Paper` и предлагает список товаров. `key=id` указывает поле идентификатора, `label=name` — поле подписи в Excel. По умолчанию поле данных содержит весь объект.

Для обратного чтения создайте файл через `renderWorkbookForm`, как в [уроке про формы](./forms.md). Чтение выбора `Paper` из такой формы вернёт:

```json
{ "product": { "id": "001", "name": "Paper" } }
```

Если нужен только идентификатор, замените тег в B1 на вариант с `return=key`:

```text
{product}{@choice:Products; key=id; label=name; return=key}
```

В `choice-report.ts` замените строку с `data` и запустите скрипт снова:

```ts
const data = { product: '001' }
```

Справочник остаётся тем же. Excel показывает `Paper`, а чтение формы возвращает `{ "product": "001" }`. Тип ключа сохраняется: строка `'001'` не становится числом.

Ключи — непустые строки или конечные числа; они должны быть уникальными. Подписи — непустые строки. При одинаковых подписях Excel показывает, например, `Paper [001]` и `Paper [002]`. Если добавление ключей не устраняет неоднозначность, рендер отклонит справочник.

`@list` и `@choice` нельзя использовать в одном поле. Обязательность и дополнительные проверки задаются отдельно через `@validate`.

## Откуда брать варианты

В `@choice` можно указать именованный справочник или путь к массиву объектов в данных:

| Источник | Где находятся варианты |
| --- | --- |
| `Products` | `options.dictionaries.Products` |
| `.availableProducts` | `availableProducts` текущей записи |
| `$root.products` | `products` в корне данных |

Например, внутри повтора `items`:

```text
{.product}{@choice:$root.products; key=id; label=name; return=key}
```

```ts
const data = {
  products: [{ id: '001', name: 'Paper' }],
  items: [{ product: '001' }],
}
```

Отчёт может брать отдельный массив `.availableProducts` из каждой записи. **Внутри повторов формы варианты должны быть общими:** именованный справочник или `$root`-массив. Индивидуальный источник записи отклоняется. Проверку допустимости выбранного варианта для конкретной записи выполняет приложение.

При чтении формы источник зависит от режима:

| Настройка | Что передать при чтении |
| --- | --- |
| `@list:Statuses` | Исходный список через `dictionaries` |
| `@choice:Products; …; return=key` | Исходный справочник через `dictionaries` |
| `@choice:$root.products; …; return=key` | Исходный массив через `context: { products }` |
| `@choice` без `return=key` | Источник уже сохранён в выданной форме; повторно передавать его не нужно |

Для `return=key` храните снимок источника при выдаче. Читатель сопоставляет подпись с переданным источником: если в свежем справочнике у той же подписи другой ID, результат может измениться. `context` содержит источники выбора и не добавляется к прочитанным данным. В объектном режиме приложение отдельно проверяет, актуален ли возвращённый объект.

## Свои сообщения

Сообщение для правила можно записать в ячейке:

```text
{quantity}{@validate:required|number|min:0}{@validationMessage:required:"Enter a quantity"}
```

Или передать сообщения для всего запуска в `options.validationMessages`:

```ts
const options = { validationMessages: { required: 'Enter a value' } }
```

Сообщение в ячейке имеет приоритет над сообщением запуска. Чтобы изменить текст только одного применения правила, укажите `message` в его директиве:

```text
{quantity}{@validate:number}{@validate:max:100; message="Enter at most 100"}
```

Такое сообщение имеет наивысший приоритет. Если переопределений нет, используется сообщение обработчика. Тексты с пробелами или знаками пунктуации заключайте в кавычки.

## Собственное правило

Новых встроенных правил из имени тега не возникает: обработчик передаёт приложение. Добавим `multipleOf`, которое требует целое число, кратное положительному целому аргументу.

Создайте лист `Input` в `rules-template.xlsx` или скачайте [готовый шаблон](/examples/tutorials/rules-template.xlsx):

| Ячейка | Содержимое |
| --- | --- |
| A1 | `Quantity` |
| B1 | `{quantity}{@validate:required\|number\|multipleOf:2}` |

После [установки](./getting-started.md) сохраните рядом [rules-report.ts](/examples/tutorials/rules-report.ts):

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'
import type { ValidationOptions } from 'sheetbind'

const options: ValidationOptions = {
  validationRules: {
    multipleOf: {
      validateArgs: args => args.length === 1 && typeof args[0] === 'number'
        && Number.isSafeInteger(args[0]) && args[0] > 0,
      validate(value, [step]) {
        return typeof value === 'number' && Number.isSafeInteger(value)
          && typeof step === 'number' && value % step === 0
      },
      message: ({ args }) => `Enter a whole number divisible by ${args[0]}`,
    },
  },
}
const template = await importWorkbookXlsx(await readFile('rules-template.xlsx'))
const data = { quantity: 4 }
await writeFile('rules-report.xlsx', await renderWorkbookReport(template, data, options))
```

Из этой папки выполните:

```sh
pnpm exec tsx rules-report.ts
```

В `rules-report.xlsx` ячейка B1 содержит `4`. Если заменить значение на `3`, рендер выбросит `TemplateError` с правилом `multipleOf` и сообщением `Enter a whole number divisible by 2`. Передача `multipleOf:0` даст ошибку настройки ещё до проверки значения.

`validateArgs` проверяет аргументы тега; без него правило не принимает аргументов. `validate` должен синхронно вернуть `true` или `false`. По умолчанию пустые значения пропускаются; для их проверки добавьте обработчику `skipEmpty: false`. Не изменяйте данные внутри обработчика и не выполняйте асинхронные запросы.

В XLSX сохраняются имя и аргументы правила. Код обработчика остаётся в приложении: передавайте одинаковые `validationRules` при выдаче и чтении формы. Встроенные правила переопределять нельзя. Типы обработчиков и ошибки исполнения описаны в [API](./api.md).

Дальше — [выдать форму, заполнить её в Excel и прочитать результат](./forms.md).
