# Learning plan

A learner has several courses, each with its own lesson list. Produce one document with headings, notes and duration totals. The courses contain zero, one and three lessons.

After [installing the package](../getting-started.md), create a directory inside your application:

```sh
mkdir study-plan-example
cd study-plan-example
```

## 1. Inspect the template

Download [study-plan-template.xlsx](/examples/walkthroughs/study-plan-template.xlsx). Its `Plan` sheet contains one course with one lesson row:

<div class="workbook-preview" tabindex="0" role="region" aria-label="Learning plan template, A1:C11">

[![The outer courses repeat covers rows 3–9. The nested lessons repeat contains row 6. A grand total follows the courses.](/images/example-study-plan-template.png)](/images/example-study-plan-template.png)

</div>

The image displays formulas as text; the file contains Excel formulas.

| Cells | Purpose |
| --- | --- |
| B1 | `{learner}` — a name from the root data |
| A2 and C10 | `{#courses}` and `{/courses}` — course boundaries |
| A3:C3 | Merged heading `{.title}` from the current course |
| A5 and C7 | `{#.lessons}` and `{/.lessons}` — the current course's lessons |
| A6:C6 | `{.name}`, `{.minutes}`, `{?.note}` — lesson fields |
| A8:B8 | `Subtotal` and `=SUBTOTAL(9,B4:B7)` — course duration |
| A9:C9 | Merged note `{.note}` from the current course |

After the inner repeat, the context is the course again: `{.note}` in A9 comes from the course, while `{?.note}` in C6 comes from a lesson. Merged cells fit entirely inside the course region and repeat with it.

## 2. Prepare the data

Save [study-plan.data.json](/examples/walkthroughs/study-plan.data.json) beside the template:

<<< @/public/examples/walkthroughs/study-plan.data.json

An empty `lessons` array removes only the lesson rows. The course name, column headings, subtotal and note remain. A missing lesson note leaves its cell blank.

## 3. Create the report

Download [study-plan.ts](/examples/walkthroughs/study-plan.ts):

<<< @/public/examples/walkthroughs/study-plan.ts

Run from the example directory:

```sh
pnpm exec tsx study-plan.ts
```

Open `study-plan.xlsx`, or download the [rendered report](/examples/walkthroughs/study-plan.xlsx):

<div class="workbook-preview" tabindex="0" role="region" aria-label="Learning plan with three courses, A1:C18">

[![Drawing has no lessons, Photography has one and Music has three. Subtotals are 0, 40 and 95, with a grand total of 135 minutes.](/images/example-study-plan.png)](/images/example-study-plan.png)

</div>

Each block retains its formatting. Lesson lists occupy different numbers of rows; their notes and the next course move accordingly.

## 4. Check totals and an empty list

The subtotals are **0, 40 and 95 minutes**, with a grand total of **135**. `SUBTOTAL(9,...)` adds numbers. Each course's range includes the text heading `Minutes`: an empty list leaves a valid range that sums to zero.

The template's B11 contains the grand total:

```text
=IFERROR(SUBTOTAL(9,B3:B10),0)
```

After rendering, it is in B18 and covers B2:B17. Excel excludes nested `SUBTOTAL` results from the grand total, so lesson durations are not counted twice. Course and lesson labels do not affect the calculation.

The image shows recalculated values. **Excel calculates formulas; Sheetbind does not.** Check the values after opening the file.

With `courses: []`, the learner name and a zero grand total remain: [empty plan](/examples/walkthroughs/study-plan-0.xlsx). You can also download [one course with no lessons](/examples/walkthroughs/study-plan-1.xlsx).

For users to fill nested lists and return them to your application, continue with [event registration](./registration.md).
