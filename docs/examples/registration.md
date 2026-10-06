# Event registration

Issue an event registration form with several groups. Each group has a name, a contact and a participant list. Users select ticket types, enter session counts and return the file; the application receives nested JSON.

First [install the package](../getting-started.md). Create a directory inside your application:

```sh
mkdir registration-example
cd registration-example
```

## 1. Inspect the template

Download [registration-template.xlsx](/examples/walkthroughs/registration-template.xlsx). The `Registration` sheet contains one group block and one participant row:

Scroll images horizontally or click them to open a larger view.

<div class="workbook-preview workbook-preview-wide" tabindex="0" role="region" aria-label="Registration template, A1:D10">

[![The outer groups repeat contains the group name, a nested participants list and a contact. Input fields have a light fill.](/images/example-registration-template.png)](/images/example-registration-template.png)

</div>

`{#groups}` in A2 and `{/groups}` in D9 enclose the whole group. Inside it, `{#.participants}` in A5 and `{/.participants}` in D7 repeat A6:D6. The group name and contact sit outside the nested list.

The ticket field in B6 is:

```text
{.ticket}
{@choice:Tickets; key=id; label=label; return=key}
{@validate:required|string}
```

All three tags occupy **the same cell B6**, each on a new line. In Excel for Windows, press Alt+Enter to insert a line break within a cell. Line breaks between tags do not change their meaning or become part of the field value.

Users see `label`; the application receives the string `id`. C6 has the binding `{.sessions}` and rule `{@validate:required|number|min:0}`: a number is required and cannot be negative. Participant names are also required. The note and contact are optional text.

## 2. Issue the form

Save [registration.data.json](/examples/walkthroughs/registration.data.json) beside the template:

<<< @/public/examples/walkthroughs/registration.data.json

Three `{}` objects create three rows for `Friends`. The empty list for `Family` creates one input row. Required fields can be empty at issuance.

Download [registration-dictionaries.json](/examples/walkthroughs/registration-dictionaries.json):

<<< @/public/examples/walkthroughs/registration-dictionaries.json

All participants in both groups share this dictionary. It is embedded in the issued form. Keep the original template for reading; the dictionary JSON is only needed when issuing the form.

Download [registration-issue.ts](/examples/walkthroughs/registration-issue.ts):

<<< @/public/examples/walkthroughs/registration-issue.ts

```sh
pnpm exec tsx registration-issue.ts
```

This creates `registration-issued.xlsx`; you can also download the [blank form](/examples/walkthroughs/registration-issued.xlsx):

<div class="workbook-preview workbook-preview-wide" tabindex="0" role="region" aria-label="Registration form with two groups and four blank rows">

[![Friends has participant rows 7–9; Family has row 17. A contact and hidden group boundaries separate the lists.](/images/example-registration-issued.png)](/images/example-registration-issued.png)

</div>

Some rows are hidden because they carry the form structure. The row numbers below refer to the issued file, not the template.

## 3. Fill and read

Open `registration-issued.xlsx` in Excel and enter:

| Row | Participant | Ticket — select from the list | Sessions | Note |
| --- | --- | --- | --- | --- |
| 7 | `Alex` | `Standard` | Number `2` | Leave blank |
| 8 | Leave the entire row blank | | | |
| 9 | `Sam` | `Student` | Number `0` | `First visit` |
| 17 | `Taylor` | `Standard` | Number `1` | Leave blank |

Keep the group names and leave contacts blank. Save a separate file named **`registration-completed.xlsx`** beside the script, or download the [completed example](/examples/walkthroughs/registration-completed.xlsx):

<div class="workbook-preview workbook-preview-wide" tabindex="0" role="region" aria-label="Completed registration with three participants">

[![Alex and Sam belong to Friends; Taylor belongs to Family. Row 9 is blank and Sam has zero sessions.](/images/example-registration-completed.png)](/images/example-registration-completed.png)

</div>

Download [registration-read.ts](/examples/walkthroughs/registration-read.ts) into the same directory:

<<< @/public/examples/walkthroughs/registration-read.ts

```sh
pnpm exec tsx registration-read.ts
```

The application receives these data, also available as a [JSON download](/examples/walkthroughs/registration-completed.json):

<<< @/public/examples/walkthroughs/registration-completed.json

Blank row 8 is omitted from `Friends`, while Sam's `0` is preserved. Tickets become `standard` and `student`; optional blank fields become `null`. A completely blank participant list is read as `[]`, while its named group remains.

## 4. Test an error and add a participant

In the original completed example, change C9 to the number `-1`, save and read again. Reading fails with a `min` issue at **C9**, path **`$data.groups[0].participants[1].sessions`**. Index `1` accounts for the omitted blank row. Download the [invalid file](/examples/walkthroughs/registration-invalid.xlsx) and [full API result](/examples/walkthroughs/registration-invalid.json) if needed. Restore C9 to `0` and save.

To add a participant to `Friends` in the completed file:

1. Select and copy the entire row 7.
2. Select row 8 and use Insert Copied Cells to insert entire rows.
3. In the new row 8, change A8 to `Morgan` and C8 to the number `1`. The ticket remains `Standard`.
4. Save the file and read it again.

`groups[0].participants` will contain `Alex`, `Morgan`, `Sam`; `groups[1].participants` will still contain `Taylor`. Inserting shifts subsequent rows: the former C9 is now C10.

Copying a whole group requires its hidden boundaries and nested list; follow the [multirow record instructions](../forms.md#multirow-records). Run validation and form reading on the application server regardless of Excel's input prompts.
