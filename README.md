
# DAILY LOGS

> A Google Apps Script sidebar for Google Sheets that combines a daily planner (TODO list with time budgets) and an activity logger with timer.

## What it does

CHRONOLOG is a container-bound Google Apps Script attached to a Google Sheet. It lets you:

- Plan your day in a TODO list with checkboxes and planned time per item.
- Log completed work with task, duration, and description.
- Compute running totals automatically per task and per day.
- Compute an Aramis ratio for each entry (based on a 7-hour reference day).
- Change the Aramis reference hours from the menu (for example 7.0, 7.5, 8.0).
- Generate a weekly recap grouped by task.
- Recalculate totals after manual edits in the sheet.

Everything runs inside Google Sheets: no external DB or backend required.

## Sidebar features

Open from `📋 Logger → Ouvrir le journal`.

### 1) TODO LIST

- Add planned tasks for the day.
- Set planned time with hour/minute selectors.
- Mark tasks done with a checkbox.
- Edit planned time via a compact badge (`hh:mm`) and pencil button.
- Use `Entrer` to prefill the logger section with the selected task.
- Remove tasks with `✕`.

TODO items are saved per-user, per-day in UserProperties with key format:

`daily_todos:yyyy-MM-dd`

### 2) Qu'as-tu fait aujourd'hui ?

- Stopwatch controls: start, pause, stop, reset.
- Duration input (`hh:mm:ss`) can be edited manually.
- Task dropdown fed by active entries in the `Tasks` sheet.
- Inline new-task creation from the sidebar.
- Optional free-text description.
- Submit writes one row into the current week sheet.

### 3) Weekly actions (in sidebar)

- `Récap semaine` button: writes/replaces the weekly summary section at the bottom of the current week sheet.
- `Recalculer totaux` button: recomputes computed columns for all valid rows in the current week.

## Spreadsheet menu

Current custom menu:

- `📋 Logger → Ouvrir le journal`
- `📋 Logger → Configurer heures Aramis`

The recap and recalculation actions are intentionally exposed in the sidebar (not in menu).

### Configure Aramis reference hours

Use `📋 Logger → Configurer heures Aramis` to set the number of hours used by Aramis calculations.

- Scope: stored per spreadsheet (`DocumentProperties`).
- Accepted values: numbers `> 0` and `≤ 24` (supports decimals like `7.5`).
- Effect: applies to new calculations immediately.
- For historical rows already written, run `Recalculer totaux` from the sidebar.

## Sheets and data model

### Weekly sheet (auto-created)

Tab name format: `WW d/MM d/MM` (ISO week based).

Columns:

1. Horodatage
2. Tâche
3. Durée
4. Durée totale tâche (running total for current day/task)
5. Description
6. Aramis
7. Aramis total tâche (running total for current day/task)
8. Durée totale jour (running total for current day)

### Tasks sheet

- Sheet name: `Tasks`
- Col A: task name
- Col B: active checkbox (`TRUE/FALSE`)

### Daily TODO storage

- API: `PropertiesService.getUserProperties()`
- Key prefix: `daily_todos:`
- Payload per day:

```json
{
  "items": [
    { "text": "Task name", "done": false, "plannedMinutes": 90 }
  ]
}
```

## Installation from git clone + clasp push

## 1) Clone the repository

```bash
git clone https://github.com/<your-user>/<your-repo>.git
cd <your-repo>
```

## 2) Install clasp

```bash
npm install -g @google/clasp
```

## 3) Authenticate with Google

```bash
clasp login
```

## 4) Bind to a Google Sheet

### Option A — Create a new bound Sheets project

```bash
clasp create --type sheets --title "CHRONOLOG" --rootDir .
```

This creates:

- a new Google Sheet
- a bound Apps Script project
- local `.clasp.json`

### Option B — Use an existing Google Sheet

1. Open the sheet.
2. Go to `Extensions → Apps Script`.
3. Copy the Script ID from URL: `https://script.google.com/d/<SCRIPT_ID>/edit`
4. Create `.clasp.json` at repo root:

```json
{
  "scriptId": "<SCRIPT_ID>",
  "rootDir": "."
}
```

## 5) Push local code to Apps Script

```bash
clasp push
```

This uploads `Code.gs` and `Sidebar.html` to the bound script project.

## 6) Open and authorize

1. Reload the Google Sheet tab.
2. Click `📋 Logger → Ouvrir le journal`.
3. Accept authorization prompts.

You can now use TODO planning + logging directly in the sidebar.

## Useful clasp commands

```bash
clasp open           # Open Apps Script editor
clasp pull           # Pull remote files locally
clasp push --watch   # Auto-push on local changes
```

## Notes

- This project is designed for Google Sheets + Apps Script (`Code.gs` + `Sidebar.html`).
- Any manual edits to durations/task names in sheet rows should be followed by `Recalculer totaux` in the sidebar.

## License

This project is intended to be distributed under the **MIT License**.

