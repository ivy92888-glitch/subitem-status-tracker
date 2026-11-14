# Configuration

All behaviour is controlled by `Private Const` values at the top of
`src/SubItemTally.bas`. Change them to match your own worksheet layout and
numbering convention, then re-run `TallySubItems`.

| Constant          | Default | Meaning                                                                 |
|-------------------|---------|-------------------------------------------------------------------------|
| `SRC_SHEET`       | `""`    | Source worksheet name. `""` means "use the currently active sheet".      |
| `SUBITEM_COL`     | `"A"`   | Column that holds the sub-item code (e.g. `PO2025001-001`).              |
| `REF_COL`         | `"D"`   | Column that holds the linked reference. **Blank = unprocessed.**         |
| `SEQ_SEPARATOR`   | `"-"`   | Character that splits `<MasterNo><Seq>` inside a sub-item code.          |
| `RESULT_SHEET`    | `"Summary"` | Name of the output worksheet (created if missing, cleared if present). |
| `ONLY_LIST_DONE`  | `True`  | `True` = list only master numbers with ≥1 processed sub-item.           |

## Tuning examples

- **Different columns:** if your sub-item code is in column `B` and the reference
  in column `F`, set `SUBITEM_COL = "B"` and `REF_COL = "F"`.
- **Different separator:** if your codes look like `LOT2025_U001` or
  `ORDER7#001`, set `SEQ_SEPARATOR = "U"` or `SEQ_SEPARATOR = "#"`.
- **Show every master number:** set `ONLY_LIST_DONE = False` to also list master
  numbers that have zero processed sub-items.
- **Fixed source sheet:** set `SRC_SHEET = "Data"` so the macro always reads that
  sheet regardless of which one is active.

## Highlight colour

Unprocessed counts are highlighted with `RGB(255, 200, 200)` (light red) in the
summary. To change it, edit the `RGB(...)` call inside the output loop.

## Grouping rule (important)

A master number is matched by **prefix** of the sub-item code. This means master
numbers must not be prefixes of one another. For example, `PO1` and `PO10` would
collide because every `PO10…` sub-item also starts with `PO1`. Use distinct,
non-overlapping master-number schemes, or pick a separator that makes the split
unambiguous for your data.
