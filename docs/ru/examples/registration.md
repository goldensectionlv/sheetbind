# Регистрация участников

Выдадим форму регистрации на мероприятие. В ней несколько групп, у каждой своё название, контакт и список участников. Пользователь выбирает тип билета, указывает число занятий и возвращает файл; приложение получает вложенный JSON.

Сначала [установите пакет](../getting-started.md). В приложении создайте папку для примера:

```sh
mkdir registration-example
cd registration-example
```

## 1. Посмотрите шаблон

Скачайте [registration-template.xlsx](/examples/walkthroughs/registration-template.xlsx). На листе `Registration` один блок группы и одна строка участника:

Картинки можно прокручивать по горизонтали и открывать крупнее нажатием.

<div class="workbook-preview workbook-preview-wide" tabindex="0" role="region" aria-label="Шаблон регистрации, A1:D10">

[![Внешний повтор groups содержит название группы, вложенный список participants и контакт. Поля ввода выделены светлой заливкой.](/images/example-registration-template.png)](/images/example-registration-template.png)

</div>

`{#groups}` в A2 и `{/groups}` в D9 охватывают всю группу. Внутри неё `{#.participants}` в A5 и `{/.participants}` в D7 повторяют строку A6:D6. Название группы и контакт находятся снаружи вложенного списка.

Тип билета в B6 задан так:

```text
{.ticket}
{@choice:Tickets; key=id; label=label; return=key}
{@validate:required|string}
```

Все три тега находятся **в одной ячейке B6**, каждый с новой строки. В Excel для Windows перенос внутри ячейки вводится через Alt+Enter. Переносы между тегами не меняют их смысл и не попадают в значение поля.

Пользователь видит `label`, а приложение получает строковый `id`. У C6 привязка `{.sessions}` и проверка `{@validate:required|number|min:0}`: число обязательно и не может быть отрицательным. Имя участника тоже обязательно. Примечание и контакт — необязательный текст.

## 2. Выдайте форму

Сохраните рядом [registration.data.json](/examples/walkthroughs/registration.data.json):

<<< @/public/examples/walkthroughs/registration.data.json

Три `{}` создадут три строки для `Friends`. Пустой список у `Family` создаст одну строку для заполнения. Пустые обязательные поля допустимы при выдаче.

Скачайте [registration-dictionaries.json](/examples/walkthroughs/registration-dictionaries.json):

<<< @/public/examples/walkthroughs/registration-dictionaries.json

Этот справочник общий для всех участников обеих групп. Он сохраняется в выданной форме. Для чтения храните исходный шаблон; JSON справочника нужен только при выдаче.

Скачайте [registration-issue.ts](/examples/walkthroughs/registration-issue.ts):

<<< @/public/examples/walkthroughs/registration-issue.ts

```sh
pnpm exec tsx registration-issue.ts
```

Получится `registration-issued.xlsx`; можно скачать [готовую пустую форму](/examples/walkthroughs/registration-issued.xlsx):

<div class="workbook-preview workbook-preview-wide" tabindex="0" role="region" aria-label="Форма регистрации с двумя группами и четырьмя пустыми строками">

[![У Friends строки участников 7–9, у Family строка 17. Между ними находятся контакт и скрытые границы групп.](/images/example-registration-issued.png)](/images/example-registration-issued.png)

</div>

Часть строк скрыта: они нужны для чтения структуры формы. Номера ниже относятся к выданному файлу, а не к шаблону.

## 3. Заполните и прочитайте

Откройте `registration-issued.xlsx` в Excel и заполните:

| Строка | Participant | Ticket — выбрать из списка | Sessions | Note |
| --- | --- | --- | --- | --- |
| 7 | `Alex` | `Standard` | Число `2` | Оставить пустым |
| 8 | Оставить всю строку пустой | | | |
| 9 | `Sam` | `Student` | Число `0` | `First visit` |
| 17 | `Taylor` | `Standard` | Число `1` | Оставить пустым |

Названия групп оставьте как есть, контакты — пустыми. Сохраните отдельный файл **`registration-completed.xlsx`** рядом со скриптом или скачайте [заполненный пример](/examples/walkthroughs/registration-completed.xlsx):

<div class="workbook-preview workbook-preview-wide" tabindex="0" role="region" aria-label="Заполненная регистрация с тремя участниками">

[![Alex и Sam относятся к Friends, Taylor — к Family. Строка 9 пустая, у Sam число занятий равно нулю.](/images/example-registration-completed.png)](/images/example-registration-completed.png)

</div>

Скачайте [registration-read.ts](/examples/walkthroughs/registration-read.ts) в ту же папку:

<<< @/public/examples/walkthroughs/registration-read.ts

```sh
pnpm exec tsx registration-read.ts
```

Приложение получит следующие данные; их можно [скачать как JSON](/examples/walkthroughs/registration-completed.json):

<<< @/public/examples/walkthroughs/registration-completed.json

Пустая строка 8 пропущена внутри `Friends`, а `0` у `Sam` сохранился. Билеты вернулись как `standard` и `student`; необязательные пустые поля — как `null`. Полностью пустой список участников читается как `[]`, сама группа с названием остаётся.

## 4. Проверьте ошибку и добавление участника

В исходном заполненном примере замените C9 на число `-1`, сохраните и повторите чтение. Оно завершится ошибкой `min`: адрес **C9**, путь **`$data.groups[0].participants[1].sessions`**. Индекс `1` учитывает пропущенную пустую строку. Готовый [ошибочный файл](/examples/walkthroughs/registration-invalid.xlsx) и [полный результат API](/examples/walkthroughs/registration-invalid.json) доступны для скачивания. Верните C9 значение `0` и сохраните файл.

Чтобы добавить участника в `Friends`, в заполненном файле:

1. Выделите и скопируйте целую строку 7.
2. Выделите строку 8 и выполните «Вставить скопированные ячейки» целыми строками.
3. В новой строке 8 измените A8 на `Morgan`, C8 — на число `1`. Билет останется `Standard`.
4. Сохраните файл и запустите чтение снова.

В `groups[0].participants` будут `Alex`, `Morgan`, `Sam`; в `groups[1].participants` останется `Taylor`. После вставки последующие строки сдвигаются: прежняя C9 теперь C10.

Копирование целой группы требует её скрытых границ и вложенного списка; порядок описан в [инструкции для многострочных записей](../forms.md#многострочные-записи). Проверки правил и чтение выполняйте на сервере приложения независимо от подсказок Excel.
