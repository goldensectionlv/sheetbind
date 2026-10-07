# Проверки и списки

Правила записываются в той же ячейке, что и [привязка](./templates.md):

```text
{quantity}
{@validate:required|number|min:0}
```

Теги одного поля можно писать с новой строки внутри одной ячейки. В Excel для Windows используйте Alt+Enter. Переносы между тегами не меняют привязку и правила; запись в одну строку тоже допустима.

Здесь `quantity` — поле данных, `required` запрещает пустое значение, `number` требует число, `min:0` запрещает отрицательные числа. Правила выполняются при чтении заполненной формы. Рендер отчёта и выдача формы не вызывают валидаторы и не требуют их обработчиков.

## Глобальная регистрация правил

Собственные правила можно один раз зарегистрировать при старте приложения:

```ts
import { registerValidationRule } from 'sheetbind'

registerValidationRule('twoLetters', {
  validate: value => typeof value === 'string' && /^[A-Z]{2}$/.test(value),
  message: 'Введите две заглавные буквы',
})
```

Ячейка `{code}{@validate:twoLetters}` использует это правило при `readWorkbookForm(template, bytes)`. Передавать обработчик в каждый вызов не нужно. `validationRules` конкретного вызова имеет приоритет над глобальной регистрацией. Встроенные имена переопределять нельзя; повторная регистрация собственного имени считается ошибкой. Регистрация действует внутри загруженного экземпляра пакета; в другом процессе или worker её нужно выполнить отдельно.

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

Если поле необязательно, не добавляйте `required`. Знак `?` в привязке разрешает отсутствие поля во входных данных отчёта; он не отменяет `required` при чтении формы. Пропущенное свойство и свойство со значением `undefined` означают отсутствие, в том числе во вложенных объектах. Заменять их на `null` перед выдачей формы не нужно.

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

Список должен содержать хотя бы одну строку. Выданная форма сохраняет этот список; при чтении справочник передавать не нужно. Значение вне списка даст ошибку `list`.

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

Отчёты и формы могут брать отдельный массив `.availableProducts` из каждой записи. Все источники списков и выбора сохраняются в выданной форме. `return=key` меняет только возвращаемое значение; повторно передавать справочник не нужно. Приложение отдельно проверяет, допустимы ли прочитанные данные сейчас.

## Разные варианты в разных строках

Поместите варианты ответа прямо в запись вопроса:

| Ячейка | Содержимое |
| --- | --- |
| A1 | `{#questions}` |
| A2 | `{.question}` |
| B2 | `{.answer}{@choice:.answers; key=id; label=label; return=key}` |
| B3 | `{/questions}` |

```ts
const data = {
  questions: [
    { question: 'Colour', answers: [{ id: 'red', label: 'Red' }, { id: 'blue', label: 'Blue' }] },
    { question: 'Size', answers: [{ id: 'small', label: 'Small' }, { id: 'large', label: 'Large' }] },
  ],
}
const issued = await renderWorkbookForm(template, data)
const result = await readWorkbookForm(template, completed)
```

Первая строка предлагает `Red / Blue`, вторая — `Small / Large`. Дополнительный идентификатор и объединение справочников не нужны. Массивы вариантов не попадают в прочитанные поля.

Индивидуальные варианты сохраняются для выданных полей. В такой форме заполняйте и очищайте значения, сохраняя состав и порядок строк. Для добавления, удаления или перестановки записей и изменения их вариантов выдайте форму заново. Автоматического переноса индивидуальных источников между строками нет.

Отсутствующие источники данных, `undefined` и `null` по путям вроде `.answers` или `$root.catalog.options` означают пустой массив. Добавлять `answers: []` в каждую запись не нужно. Именованные справочники по-прежнему нужно передавать через `options.dictionaries`.

Если при пустом или отсутствующем `answers` допустим текст, добавьте `emptySource=input` вместе с `return=key`. Непустой источник по-прежнему требует выбрать его вариант, а правила `@validate` продолжают работать. При пустом источнике поле допускает обычный ввод без выпадающего списка.

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

После [установки](./getting-started.md) сохраните рядом [rules-form.ts](/examples/tutorials/rules-form.ts):

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookForm, readWorkbookForm } from 'sheetbind'
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
await writeFile('rules-form.xlsx', await renderWorkbookForm(template, data))
const result = await readWorkbookForm(template, await readFile('rules-form.xlsx'), options)
console.log(JSON.stringify(result, null, 2))
```

Из этой папки выполните:

```sh
pnpm exec tsx rules-form.ts
```

Скрипт выдаёт `rules-form.xlsx` и читает сохранённый файл. При `quantity: 4` чтение успешно. Если заменить значение на `3`, файл всё равно создаётся, а чтение возвращает `success: false` с правилом `multipleOf` и сообщением `Enter a whole number divisible by 2`. `multipleOf:0` вызывает ошибку настройки при чтении, до проверки значения.

`validateArgs` проверяет аргументы тега; без него правило не принимает аргументов. `validate` должен синхронно вернуть `true` или `false`. По умолчанию пустые значения пропускаются; для их проверки добавьте обработчику `skipEmpty: false`. Не изменяйте данные внутри обработчика и не выполняйте асинхронные запросы.

В XLSX сохраняются имя и аргументы правила. Код обработчика остаётся в приложении: передавайте `validationRules` только при чтении формы. Неизвестное правило останавливает чтение с ошибкой шаблона; передайте его обработчик до чтения. Встроенные правила переопределять нельзя. Типы обработчиков и ошибки исполнения описаны в [API](./api.md).

Дальше — [выдать форму, заполнить её в Excel и прочитать результат](./forms.md).

## Форматирование значений {#formatting}

Форматтеры преобразуют значение при рендере отчёта, выдаче формы и вызове `resolveWorkbook`. Исходный объект данных не меняется. Форматирование задаётся в ячейке:

```text
{deliveryDate | format_date:DD.MM.YYYY}
{enabled | bool_replace:"Да","Нет"}
{price | float}
```

Тот же форматтер можно вынести на отдельную строку внутри ячейки:

```text
{deliveryDate}
{@format:format_date:DD.MM.YYYY}
```

| Форматтер | Результат |
| --- | --- |
| `float` | Числовая строка становится конечным числом. Пустые и непригодные значения сохраняются |
| `bool_replace:"Да","Нет"` | Первая подпись для истинного значения, вторая для ложного |
| `format_date:DD.MM.YYYY` | Строка даты или Unix timestamp в миллисекундах становится текстом по маске. По умолчанию — `DD.MM.YYYY`; используются локальные компоненты даты процесса |

Маска даты поддерживает `YYYY`, `YY`, `MM`, `M`, `DD`, `D`, `HH`, `H`, `mm`, `m`, `ss`, `s`. Маску со временем заключайте в кавычки: `format_date:"DD.MM.YYYY HH:mm"`. Нераспознаваемая дата сохраняет исходное значение. Для количества знаков после запятой у уже числовой ячейки используйте обычный формат Excel, например `0.00`.

Свой форматтер зарегистрируйте один раз при старте приложения:

```ts
import { registerFormatter } from 'sheetbind'

registerFormatter('uppercase', value => value === null ? null : String(value).toUpperCase())
registerFormatter('suffix', (value, args) => String(value ?? '') + String(args[0] ?? ''))
```

```text
{code | uppercase | suffix:"!"}
```

Встроенные форматтеры сохраняют `null`, пустую строку и пробельный текст. Пустой ответ не становится `0` или отрицательной подписью; настоящие `0` и `false` продолжают обрабатываться.

Цепочка выполняется слева направо. Обработчик получает значение и массив JSON-аргументов; возвращает строку, конечное число, boolean или `null` синхронно. Незнакомый форматтер, ошибка обработчика или недопустимый результат останавливают рендер. Встроенные имена защищены, повторная регистрация своего имени считается ошибкой.

Форматтеры **не выполняются при чтении** и не обращают преобразование: после `bool_replace` читается текст «Да», а не boolean. Для возврата ключа или объекта по подписи используйте `@choice`. Форматирование нельзя совмещать с `@choice` или `@list` в одной ячейке. Регистрация обработчиков не записывается в XLSX.

[Полный запускаемый пример](/examples/tutorials/formatting.ts) создаёт форму со встроенными и своим форматтерами и читает её с глобальным правилом. После установки зависимостей выполните `pnpm exec tsx formatting.ts`.
