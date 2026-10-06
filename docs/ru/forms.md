# Заполнение формы

Форма — это XLSX, который пользователь заполняет и возвращает приложению. `renderWorkbookForm` создаёт такой файл, `readWorkbookForm` читает его по исходному шаблону. Для выполнения примеров сначала [установите пакет и настройте запуск TypeScript](./getting-started.md).

## Первая форма

### 1. Создайте шаблон

Создайте лист `Input` в `first-form-template.xlsx` или скачайте [готовый шаблон](/examples/tutorials/first-form-template.xlsx):

| Ячейка | Содержимое |
| --- | --- |
| A1 | `Customer` |
| B1 | `{customer.name}{@validate:required\|string}` |

Поле принимает непустой текст. Обычные подписи, например `Customer`, в результат чтения не попадут.

### 2. Выдайте файл

Сохраните [first-form.ts](/examples/tutorials/first-form.ts) рядом с шаблоном:

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookForm } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('first-form-template.xlsx'))
const data = { customer: { name: 'Sample customer' } }
await writeFile('first-form-issued.xlsx', await renderWorkbookForm(template, data))
```

Выполните из этой папки:

```sh
pnpm exec tsx first-form.ts
```

Откройте `first-form-issued.xlsx`. Поле находится в B2: форма добавила скрытую служебную строку перед содержимым. Измените имя на `Alex` и сохраните **отдельный файл `completed.xlsx`** рядом со скриптом. Исходный шаблон сохраните для чтения.

### 3. Прочитайте заполненный файл

Сохраните [read-first-form.ts](/examples/tutorials/read-first-form.ts) в той же папке:

```ts
import { readFile } from 'node:fs/promises'
import { importWorkbookXlsx, readWorkbookForm } from 'sheetbind'

try {
  const template = await importWorkbookXlsx(await readFile('first-form-template.xlsx'))
  const result = await readWorkbookForm(template, await readFile('completed.xlsx'))
  if (result.success) {
    console.log(JSON.stringify(result.data, null, 2))
  }
  else {
    console.error(result.issues)
    process.exitCode = 1
  }
}
catch (error) {
  console.error('Cannot process the form:', error)
  process.exitCode = 1
}
```

```sh
pnpm exec tsx read-first-form.ts
```

Результат:

```json
{
  "customer": { "name": "Alex" }
}
```

Если очистить B2 и снова сохранить файл, `success` будет `false`. В `issues` появится ошибка `required` с `sheetName: "Input"`, `address: "B2"`, `path: "$data.customer.name"`. При ошибках частичных `data` нет.

`issues` описывает проблемы возвращённого файла или введённых значений. Неверные настройки приложения и ошибки работы с файлами могут выбрасывать исключения, поэтому нужны и `result.success`, и `try/catch`. Подробности — в [API](./api.md).

## Таблица на 20 строк

Теперь выдадим пустые строки со списком товаров, заполним две и оставим промежуток. В примере используются [проверки и списки](./fields.md): `required` требует значение, `number|min:0` — неотрицательное число, `@choice` показывает товар и возвращает его ID.

### 1. Подготовьте шаблон

В папке приложения из первого урока создайте подпапку для этого примера:

```sh
mkdir table-form
cd table-form
```

Она использует уже установленные пакеты и настройку модулей приложения. Следующие файлы сохраняйте и команды выполняйте в `table-form`.

Создайте лист `Input` в `form-template.xlsx` или скачайте [шаблон](/examples/tutorials/form-template.xlsx):

| Ячейка | Содержимое |
| --- | --- |
| A1:C1 | Заголовки `Product`, `Quantity`, `Date` |
| A2 | `{#items}` |
| A3 | `{.product}{@choice:Products; key=id; label=name; return=key}{@validate:required\|string}` |
| B3 | `{.quantity}{@validate:required\|number\|min:0}` |
| C3 | `{?.date}{@validate:string}` |
| C4 | `{/items}` |

Повторяется строка A3:C3. Оформите её и задайте B3 числовой формат `0.00`. Для C3 нужен текстовый формат `@`: дату вводим строкой `2026-01-15`. Нативные даты Excel в полях ввода не поддерживаются; правило `string` проверяет тип, а не корректность календарной даты.

### 2. Выдайте строки

Сохраните [issue-form.ts](/examples/tutorials/issue-form.ts) рядом с шаблоном:

```ts
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookForm } from 'sheetbind'

const source = await readFile('form-template.xlsx')
const template = await importWorkbookXlsx(source)
const dictionaries = {
  Products: [{ id: '001', name: 'Paper' }, { id: '002', name: 'Paper' }],
}
const data = { items: Array.from({ length: 20 }, () => ({})) }
const issued = await renderWorkbookForm(template, data, { dictionaries })

await mkdir('saved-form', { recursive: true })
await writeFile('saved-form/template.xlsx', source)
await writeFile('saved-form/dictionaries.json', JSON.stringify(dictionaries, null, 2))
await writeFile('saved-form/issued.xlsx', issued)
```

```sh
pnpm exec tsx issue-form.ts
```

Двадцать объектов `{}` создают двадцать строк для заполнения. Пустые обязательные поля допустимы при выдаче. У товаров одинаковые подписи, поэтому список показывает `Paper [001]` и `Paper [002]`.

Так выглядит верхняя часть [выданной формы](/examples/tutorials/issued.xlsx). Строки ввода начинаются с 4 и продолжаются до 23; служебные строки 1 и 3 скрыты.

<div class="workbook-preview" tabindex="0" role="region" aria-label="Пустая форма, колонки A–C, начало таблицы">

[![Выданная форма: заголовки Product, Quantity, Date на строке 2, пустые строки ввода с 4. В колонке A доступны списки товаров.](/images/form-issued.png)](/images/form-issued.png)

</div>

В этом примере используется `return=key`. Храните **справочник выдачи** вместе с шаблоном: при чтении подпись преобразуется в ID по переданному справочнику. Свежие данные API могут дать другое соответствие. Объектный режим и другие источники описаны в [проверках и списках](./fields.md).

### 3. Заполните и прочитайте

Откройте `saved-form/issued.xlsx`. Строки ввода находятся на строках 4–23:

| Строка | Product | Quantity | Date |
| --- | --- | --- | --- |
| 4 | Выберите `Paper [001]` | Число `2` | Текст `2026-01-15` |
| 5 | Оставьте пустой | | |
| 6 | Выберите `Paper [002]` | Число `0` | Оставьте пустой |

После заполнения двух строк получится:

<div class="workbook-preview" tabindex="0" role="region" aria-label="Заполненная форма, колонки A–C, строки 4–6">

[![В строке 4 выбраны Paper [001], количество 2.00 и дата 2026-01-15. Строка 5 пустая. В строке 6 выбраны Paper [002] и количество 0.00, дата пустая.](/images/form-completed.png)](/images/form-completed.png)

</div>

Сохраните файл как `completed.xlsx` рядом с `issue-form.ts` или скачайте [готовую заполненную форму](/examples/tutorials/completed.xlsx). Для чтения нужны созданные на предыдущем шаге `saved-form/template.xlsx` и `saved-form/dictionaries.json`. Скачайте рядом с `completed.xlsx` скрипт [read-completed-form.ts](/examples/tutorials/read-completed-form.ts) и запустите:

```sh
pnpm exec tsx read-completed-form.ts
```

Это полный обработчик с проверкой `result.success` и `try/catch`, как в первой форме. Он загружает `saved-form/template.xlsx`, `saved-form/dictionaries.json` и `completed.xlsx` из текущей папки.

Результат можно [скачать как JSON](/examples/tutorials/completed.json):

<<< @/public/examples/tutorials/completed.json

Пустая строка 5 и остальные незаполненные строки не вошли в массив; значение `0` сохранилось. Адрес ячейки относится к Excel, индекс в данных — к массиву после удаления пустых строк.

Чтобы проверить ошибку чтения, скачайте [invalid.xlsx](/examples/tutorials/invalid.xlsx) и сохраните в папке примера под именем `completed.xlsx`. В нём A6 уже содержит `Unknown`. Обычный ввод такого значения Excel блокирует проверкой списка, поэтому здесь используется готовый файл. Повторный запуск вернёт `success: false` и ошибку `choice` с адресом `A6` и путём `$data.items[1].product`; [полный результат](/examples/tutorials/invalid.json) можно скачать. Выберите в A6 `Paper [002]`, сохраните файл и прочитайте его снова.

## Пустые строки и значения

Для однострочного повтора с полями ввода `items: []` создаёт одну пустую строку. Чтобы получить заданное количество строк, передайте столько объектов `{}`, как в примере выше.

При чтении полностью пустые однострочные записи пропускаются. Частично заполненные проходят проверку; строка с непустым ошибочным значением тоже не исчезает. `null`, пустой текст и текст из пробелов считаются пустыми; `0` и `false` — заполненными. Пустая ячейка в сохранённой записи возвращается как `null`.

Типы не преобразуются: текст `"2"` остаётся строкой и не проходит `number`. Формула в поле ввода даёт ошибку `formula`, даже если Excel сохранил её результат. Расчётные ячейки без привязки не входят в прочитанные данные.

## Добавление, удаление и перестановка записей

### Однострочные записи

В таблице выше копируйте целую строку Excel и вставляйте копию внутри области ввода. Для удаления удалите целую строку; для перестановки вырежьте и вставьте целую строку в той же области. Копирование сохраняет оформление, формулы и списки. Очистка всех полей однострочной записи исключает её из результата.

Формы поддерживают повторы вниз. Повтор вправо предназначен для отчётов. Располагайте независимые списки формы один под другим: их диапазоны строк не должны пересекаться. Вложенные списки поддерживаются.

Не удаляйте скрытые служебные строки и столбцы отдельно от записей и не переименовывайте листы с полями.

### Многострочные записи

Если одна запись занимает несколько строк, копируйте весь блок со скрытыми границами. Для примера скачайте [contacts-template.xlsx](/examples/tutorials/contacts-template.xlsx) и [выданную форму с двумя записями](/examples/tutorials/contacts-issued.xlsx). Разметка листа `Contacts`:

| Ячейка | Содержимое |
| --- | --- |
| A1 | `{#contacts}` |
| A2 | `Name` |
| B2 | `{.name}` |
| A3 | `Email` |
| B3 | `{.email}` |
| B4 | `{/contacts}` |

В выданной форме первая запись занимает целые строки **3:6**, вторая — **7:10**. Поля видны на строках 4:5 и 8:9; остальные строки блока скрыты.

1. В поле имени слева от строки формул Excel введите `3:6` и нажмите Enter. Выделятся целые строки со скрытыми границами.
2. Скопируйте их. В поле имени введите `7:7` и нажмите Enter, чтобы выбрать скрытую строку 7. Выполните «Вставить скопированные ячейки» целыми строками.

После этого одного копирования сохраните книгу как `contacts-completed.xlsx` рядом с `contacts-template.xlsx`. Скачайте туда [read-contacts.ts](/examples/tutorials/read-contacts.ts) и запустите:

```sh
pnpm exec tsx read-contacts.ts
```

Получатся три записи:

```json
{
  "contacts": [
    { "name": "Alex", "email": "alex@example.com" },
    { "name": "Alex", "email": "alex@example.com" },
    { "name": "Sam", "email": "sam@example.com" }
  ]
}
```

Удаление и перемещение — отдельные действия:

- Для удаления выделите весь блок и выполните «Удалить строки листа». Клавиша Delete только очищает ячейки.
- Для перемещения вырежьте весь блок и вставьте вырезанные строки перед границей другого блока.

Номера относятся к исходному примеру: после вставки следующие блоки сдвигаются. Не копируйте только видимые поля. **Очистка многострочного блока сохраняет пустую запись**, которая может не пройти `required`; удаление блока удаляет запись. При копировании родителя с вложенным списком включайте весь список и границы родителя.

После изменения привязок, правил или структуры шаблона выдавайте формы заново. Для чтения сохраняйте исходный шаблон и используемые обработчики правил. Сравнение с прежними данными и разрешённые изменения определяет ваше приложение. Остальные особенности файлов описаны в [Excel и ограничениях](./xlsx.md).
