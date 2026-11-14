# subitem-status-tracker

A small, dependency-free **WPS / Excel VBA** macro that groups sub-items
(boxes, batches, line items, serials, …) under a **master number** and tallies
how many are **processed** vs. **unprocessed**, then builds a one-click summary
sheet — with highlighting, borders, auto-filter and a frozen header.

It started as an internal bulk-inventory tally tool and has been rewritten into a
generic, reusable utility that works for many roles: warehouse picking, logistics
manifest reconciliation, production batch tracking, order-fulfilment checks,
inventory QC, and more.

> *Unofficial project. Not affiliated with, endorsed by, or sponsored by Kingsoft
> (WPS Office) or Microsoft (VBA). Product names are used only to describe the
> runtime environment.*

## Features

- **Group & count** — automatically derives the master number from each sub-item
  code and counts total / processed / unprocessed per master number.
- **One-click report** — writes (or refreshes) a `Summary` worksheet with a clean
  table, bold header, auto-fitted columns, borders, auto-filter and a frozen
  first row.
- **Visual alert** — unprocessed counts are highlighted in light red so exceptions
  pop out.
- **Configurable** — columns, separator, sheet names and output rules are plain
  constants at the top of the module (see [`docs/Configuration.md`](docs/Configuration.md)).
- **No add-ins** — pure VBA using the built-in `Scripting.Dictionary`; runs in
  WPS Office and Microsoft Excel on Windows.

## How it works

1. Read every data row from the source sheet (sub-item code in one column, a
   linked reference in another).
2. Derive the **master number** by stripping the sequence suffix from the
   sub-item code (everything before the configured separator).
3. Build two dictionaries: the set of master numbers, and the processed /
   unprocessed status of each sub-item.
4. For each master number, sum total, processed and unprocessed sub-items, and
   write one row to the summary sheet.
5. Format the result and report completion.

> **Note on grouping:** a master number is matched by *prefix* of the sub-item
> code. Master numbers must not be prefixes of one another (e.g. `PO1` and
> `PO10` would collide). Use distinct, non-overlapping master-number schemes, or
> adjust the separator so the split is unambiguous.

## Requirements

- WPS Office (with VBA support) **or** Microsoft Excel on Windows.
- A data sheet where:
  - one column holds the **sub-item code** (e.g. `PO2025001-001`), and
  - another column holds a **linked reference** that is present when the sub-item
    is processed and **blank** when it is not.

## Installation

1. Open your workbook in WPS / Excel.
2. Press `Alt + F11` to open the VBA editor.
3. `File ▸ Import File…` and select [`src/SubItemTally.bas`](src/SubItemTally.bas).
4. Close the editor.

## Usage

1. Put your data on a sheet (sub-item code in column **A**, linked reference in
   column **D** by default; headers in row 1, data from row 2).
2. Open the Macros dialog (`Alt + F8`), choose **`TallySubItems`**, and run it.
3. A `Summary` sheet is created (or refreshed) with the tallies.

To adapt the columns, separator, sheet names or output rule, edit the constants
at the top of the module — see [`docs/Configuration.md`](docs/Configuration.md).

## Sample data

[`examples/sample-data.xlsx`](examples/sample-data.xlsx) contains neutral example
data you can open and run the macro against immediately. Expected result:

| MasterNo   | Total | Processed | Unprocessed |
|------------|-------|-----------|-------------|
| PO2025001  | 5     | 2         | 3           |
| PO2025002  | 2     | 2         | 0           |

(`PO2025003` has only unprocessed sub-items, so with the default
`ONLY_LIST_DONE = True` it is omitted from the summary.)

## Project layout

```
subitem-status-tracker/
├── src/
│   └── SubItemTally.bas        # the macro (import into WPS/Excel)
├── docs/
│   └── Configuration.md        # how to tune the constants
├── examples/
│   └── sample-data.xlsx        # neutral demo data
├── LICENSE                     # MIT
├── README.md
├── .gitignore
├── .gitattributes
└── PUSH_GUIDE.md               # how this repo was published to GitHub
```

## Compatibility

Tested logic in WPS Office on Windows. The macro uses late-bound
`Scripting.Dictionary` and standard Excel object model calls, so it also runs in
desktop Microsoft Excel for Windows.

## License

Released under the [MIT License](LICENSE). Copyright (c) 2026 ivy92888-glitch.

## Legal

- This project is an independent, clean-room reimplementation authored by
  `ivy92888-glitch`. It contains **no** proprietary data, screenshots, or
  business-specific material from any former employer or third party.
- The MIT License covers only the author's own implementation. Publishing it does
  not represent or transfer any rights of any other party.
- Product names such as WPS Office and Microsoft Excel are trademarks of their
  respective owners and are referenced solely to describe the runtime
  environment.
