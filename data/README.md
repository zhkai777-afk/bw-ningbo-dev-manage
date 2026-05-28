# Gate Checklist Task Data

This directory stores the task data split from the first worksheet `Gate Checklist-更新版` in the source Excel file:

`/Users/zhazhakai777/Desktop/副本Gate Checklist _ Auros (002) 的副本_副本.xls`

## Split Rules

- Only the first updated worksheet was used.
- One checklist row is expanded into one task per marked stage.
- `For all gates` tasks are expanded to every source stage.
- `X` is treated as `required`.
- `√` is treated as `optional`.
- When the description contains stage-prefixed text such as `Gate 4: ...`, the matching stage gets that specific description.
- When no stage-specific text exists, the full description is reused for the stage task.

## Files

- `gate_tasks_split.raw.json`: Full source-stage split data, including `Gate 0`, `CKO`, `PKO`, and `Gate 1` to `Gate 6`. Total: 345 records.
- `gate_tasks_split.raw.csv`: CSV version of the raw split data.
- `gate_tasks_summary_by_stage.raw.csv`: Raw source-stage summary.
- `gate_tasks_split.ui.json`: UI-ready split data for this project. `Gate 0` tasks are merged into `CKO`, duplicates are removed, and `uiStage` is added. Total: 333 records.
- `gate_tasks_split.ui.csv`: CSV version of the UI-ready data.
- `gate_tasks_summary_by_ui_stage.csv`: UI-stage summary.

## UI Stage Counts

- CKO: 37
- PKO: 44
- Gate1: 39
- Gate2: 50
- Gate3: 39
- Gate4: 62
- Gate5: 43
- Gate6: 19

Use `gate_tasks_split.ui.json` for pages that follow the current project UI stages. Use `gate_tasks_split.raw.json` if a future page needs to preserve the original Excel stages exactly.
