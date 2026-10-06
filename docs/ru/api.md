# API

Импортируйте функции и типы из `sheetbind`. Буфер Node.js `Buffer` подходит везде, где функция принимает `Uint8Array`. Теги XLSX описаны в [разметке шаблона](./templates.md), проверки и выбор значений — в [правилах полей](./fields.md).

## Данные и параметры

Данные запуска должны быть обычным JSON-объектом. Внутри допустимы объекты, массивы, строки, конечные числа, boolean и `null`. Отсутствующее свойство пропустите или задайте как `null`; `undefined`, `Date`, функции, геттеры, циклические объекты и разреженные массивы отклоняются. Экспортируемый тип `JsonValue` описывает JSON-значения.

Ниже используются два сокращения. **`RunOptions` и `ReadOptions` не экспортируются из пакета**; они обозначают такую форму параметров:

```ts
type RunOptions = ValidationOptions & {
  readonly dictionaries?: Dictionaries
}

type ReadOptions = RunOptions & {
  readonly context?: Readonly<Record<string, unknown>>
}
```

| Параметр | Где применяется | Назначение |
| --- | --- | --- |
| `dictionaries` | Размещение, рендер, чтение | Именованные списки строк или объектов, объявленные в полях |
| `validationRules` | Размещение, рендер, чтение | Собственные синхронные правила полей |
| `validationMessages` | Размещение, рендер, чтение | Тексты ошибок по имени правила |
| `context` | Только чтение | Общие коллекции выбора, доступные через `$root`; эти значения не попадают в отправленные данные |

`Dictionaries` — объект только для чтения, содержащий массивы строк или JSON-объектов. `ValidationOptions` содержит `validationRules` и `validationMessages`. Собственный обработчик имеет тип `ValidationRule`; его предикат получает `ValidationContext` (`root`, `current`, `path`). `ValidationMessage` — строка или функция с аргументом `ValidationMessageContext`, в котором также есть `value`, `rule`, `args` и `index`. См. [полный пример обработчика](./fields.md).

## Загрузка шаблона

```ts
declare function importWorkbookXlsx(bytes: Uint8Array): Promise<WorkbookTemplate>
```

Загружает размеченный XLSX и проверяет теги и геометрию повторов. Данные приложения и предикаты полей на этом этапе не проверяются.

`WorkbookTemplate` — тип результата загрузки. Публичного конструктора и методов редактирования у него нет. Результат можно использовать для нескольких запусков; для хранения оставляйте исходный XLSX и загружайте его заново после изменения. Для аннотации используйте `import type { WorkbookTemplate } from 'sheetbind'`.

## Рендер отчёта

```ts
declare function renderWorkbookReport(
  template: WorkbookTemplate,
  data: unknown,
  options?: RunOptions,
): Promise<Buffer>
```

Подставляет значения, разворачивает повторы, проверяет правила полей и возвращает байты отчёта. Неверные значения полей останавливают рендер. Сохраните буфер в файл `.xlsx`. Пересчёт формул и сохранение возможностей Excel описаны в [Excel и ограничениях](./xlsx.md).

## Выдача формы

```ts
declare function renderWorkbookForm(
  template: WorkbookTemplate,
  data: unknown,
  options?: RunOptions,
): Promise<Buffer>
```

Возвращает редактируемый XLSX со структурой для последующего чтения. Проверяет шаблон и настройку правил; обязательные поля при выдаче могут быть пустыми. Значения полей проверяются при чтении заполненной формы.

Пустой массив для повтора из одной строки полей создаёт одну пустую строку ввода. Для 20 строк передайте массив из 20 пустых объектов. Формы поддерживают повтор вниз по листу; допустимые изменения и многострочные записи описаны в [работе с формами](./forms.md).

## Чтение заполненной формы

```ts
declare function readWorkbookForm(
  template: WorkbookTemplate,
  bytes: Uint8Array,
  options?: ReadOptions,
): Promise<WorkbookFormResult>
```

Передайте исходный шаблон и заполненный файл. Чтение возвращает поля, объявленные привязками, после разрешения выбора и исключения пустых однострочных записей. Результат не объединяется с данными выдачи; сопоставления записей по идентификатору нет.

```ts
type WorkbookFormResult =
  | { readonly success: true, readonly data: Readonly<Record<string, unknown>> }
  | { readonly success: false, readonly issues: readonly WorkbookFormIssue[] }
```

При ошибке частичные `data` не возвращаются. Проверяйте `result.success` и отдельно обрабатывайте исключения из настройки шаблона или приложения.

Для строковых списков и выбора с возвратом ключа передайте справочники выдачи повторно. Корневые коллекции для выбора ключа передаются в `context`. При возврате объекта используется источник, встроенный в выданный файл. В [руководстве по формам](./forms.md) показано, что сохранить для чтения.

## Просмотр значений и размещения

```ts
declare function resolveWorkbook(
  template: WorkbookTemplate,
  data: unknown,
  options?: RunOptions,
): WorkbookLayout
```

Синхронно подставляет и проверяет данные отчёта, не создавая файл. Возвращает размещение отчёта; при выдаче формы добавляются скрытые строки для её структуры.

`WorkbookLayout.sheets` содержит листы с `id`, `name`, необязательными `state`, `rows`, `columns`, `print` и массивом `cells`. Каждую ячейку описывает `WorkbookCellInstance`:

| Поле | Значение |
| --- | --- |
| `id`, `definitionId` | ID ячейки результата и исходной импортированной ячейки |
| `at` | `GridAddress`: `{ row, column }`, отсчёт с 1 |
| `size` | `GridOffset`: `{ rows, columns }`, включая размер объединённой ячейки |
| `value` | `{ literal: TemplateValue }` или `{ formula: string }`; формула без начального `=` |
| `rules` | Необязательные `FieldRules`: проверки, сообщения, строковый список или выбор |
| `origin` | `{ nodeId, dataPath, iterations: [{ nodeId, index }] }`: исходный узел, путь данных и индексы повторов |
| `contextPath` | Путь данных контекста привязки |
| `choice` | Необязательный объект `{ text, items: [{ key, label, value, text }] }`: текст выбора и варианты |

`TemplateValue` — `string | number | boolean | null`. `FieldValue` также допускает объект выбора. `ChoiceOption` описывает вариант без текста отображения в Excel; `ResolvedChoice` — выбранный ключ и такие варианты. `ValueExpression` описывает литерал или привязку к данным; привязка использует `DataReference` (`path`, необязательный `from: 'current' | 'root'`). Эти типы не предоставляют API создания шаблона.

Результат доступен только для чтения в TypeScript. Ячейки и настройки листов скопированы: изменения результата из JavaScript не затронут соседние ячейки, шаблон, входные данные и следующие запуски. ID относятся к импортированному шаблону и исполнению; это не идентификаторы записей данных. Части исходного XLSX и выражения тегов в размещение не входят.

## Список нужных справочников

```ts
declare function workbookDictionarySources(template: WorkbookTemplate): string[]
```

Возвращает отсортированные уникальные имена строковых списков и именованных источников выбора. Не подставляет данные и не включает коллекции контекста вроде `$root.products`. Приложение загружает справочники и передаёт их через `options.dictionaries`.

## Разбор и запись правил

```ts
declare function parseValidation(value: unknown): readonly ValidationRuleUse[]
declare function formatValidation(value: Validation): string
```

`Validation` — строка правил или упорядоченный массив `ValidationRuleUse`. Каждое применение содержит имя `rule`, необязательные JSON-аргументы `args` и необязательное строковое `message`.

```ts
import { parseValidation, formatValidation } from 'sheetbind'

const rules = parseValidation('required|number|min:0')
const text = formatValidation(rules) // 'required|number|min:0'
```

Разбор проверяет синтаксис и нормализует список; правила не выполняются, наличие собственного обработчика не проверяется. Имена обработчиков и аргументы проверяются при подготовке исполнения, в том числе для пустых повторов. Запись в строку исключает сообщения отдельных применений; для их сохранения используйте массив. `ValidationIssue` описывает ошибку правила: `code`, `rule`, `args`, `index`, `message`.

## Имя диапазона выбора

```ts
declare function workbookChoiceRange(rule: ChoiceRule, field?: string): string
```

Возвращает имя диапазона Excel для **именованного выбора объектов**. Без `field` это диапазон отображаемых вариантов; с `field` — значений свойства верхнего уровня. Имя зависит от названия справочника и настроек `key` и `label`.

```ts
import { workbookChoiceRange } from 'sheetbind'
import type { ChoiceRule } from 'sheetbind'

const rule: ChoiceRule = {
  source: { dictionary: 'Products' },
  key: 'id',
  label: 'name',
}
const labels = workbookChoiceRange(rule)
const prices = workbookChoiceRange(rule, 'price')
const formula = `INDEX(${prices},MATCH(B3,${labels},0))`
```

Полученную формулу можно записать в исходную книгу. Рендер создаёт диапазоны, когда в результат попадает поле выбора объектов с этим справочником (`return` пропущен или равен `'object'`). Функция лишь вычисляет имя, а не создаёт диапазон. Диапазон свойства существует, только если это свойство есть в объектах источника. Контекстный источник вызывает `RangeError`; выбор с возвратом ключа не создаёт эти диапазоны свойств.

`ChoiceRule` содержит `source`, `key`, `label` и необязательный `return: 'object' | 'key'`. Источник имеет вид `{ dictionary: string }` или `DataReference`.

## Ошибки

| Ситуация | Результат |
| --- | --- |
| Неверные значения или структура возвращённой формы | `WorkbookFormResult` с `success: false` |
| Возвращённый файл не читается как XLSX | Такой же результат с `code: 'invalid-workbook'` |
| Нечитаемый шаблон или неверный тег XLSX | `TaggedXlsxError` |
| Не хватает данных или справочников, неверные правила поля | `TemplateError`, часто его подкласс `TaggedXlsxError` |
| Собственный предикат или функция сообщения упали либо вернули неверный тип | `ValidationExecutionError` |
| Неверные аргументы, неподдерживаемая геометрия или конфликт нативной валидации | `SyntaxError`, `TypeError` или `RangeError` в зависимости от проверки |

`TemplateError.issues` содержит объекты `TemplateIssue` с `phase: 'template' | 'data'`, `code`, `path`, `nodeId` и `message`. Ошибки правил могут дополнительно содержать `rule`, `args` и `index`.

`TaggedXlsxError` наследует `TemplateError`. Его `TaggedXlsxIssue` добавляет необязательные `sheetName` и `address` исходного тега. Ошибка нечитаемого шаблона сохраняет исходную причину в `cause`.

`ValidationExecutionError` содержит `rule`, `path` и `cause`. Предикат должен синхронно вернуть boolean, функция сообщения — строку. Promise недопустим в обоих случаях.

`WorkbookFormIssue` содержит `phase: 'structure' | 'value' | 'xlsx'`, `code`, `path`, `message` и необязательные `sheetName`, `address`, `nodeId`, `rule`, `args`, `index`. Адрес относится к возвращённой книге, а путь данных использует индексы после исключения пустых строк. Для некоторых ошибок структуры определить ячейку невозможно. В [примере чтения формы](./forms.md) показана обработка и `try/catch`, и `result.success`.
