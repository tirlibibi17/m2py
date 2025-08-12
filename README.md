# M → Python (Power Query to pandas)

A scrappy, “vacation-project” converter that takes **Power Query (M)** and emits **pandas** code. You can paste M, or (on Windows) read queries straight from an Excel workbook via COM. It’s not a full M parser—just enough to be useful for quick migrations and prototyping.

> ⚠️ Alpha quality. Expect rough edges. Please file issues with examples!

---

## What it does

- **Streamlit app** with two tabs:
  - **Paste M** → convert to pandas.
  - **Excel (Windows/COM)** → upload a workbook; auto-loads queries; auto-converts the selected one and its dependencies.
- **Dependency resolution**: if a query references another (e.g., `#"Some Query"`), the bundle includes the dependency first.
- **`Excel.CurrentWorkbook` data**: the app pre-builds a `__cw` dict with real `pd.DataFrame`s pulled from the uploaded workbook so code runs immediately.
- **Commented external-Excel helper**: the bundle includes a commented snippet showing how to read the same tables from a normal `.xlsx` file with `openpyxl`.

### Supported patterns (non-exhaustive)

- `Table.FromRecords`, `Table.FromRows`, `#table`
- `Csv.Document(File.Contents(...))` → `pd.read_csv(...)` equivalent
- `Table.PromoteHeaders`, `Table.TransformColumnTypes`
- `Table.SelectRows` (simple filters), `Table.Sort`
- `Table.Join` (basic joins, single/multi-key)
- `Table.Group` with aggregations:
  - Sum, Average, Count, Min, Max, Median, StdDev, Variance, First, Last, Product
- Direct query references: `Source = #"Other Query"` or `Source = Other_Query`

If something isn’t recognized, the converter leaves a clear `# Unsupported:` comment and a safe no-op placeholder, so the Python still runs.

---

## Quickstart

### 1) Install

# Python 3.10+ recommended
`pip install streamlit pandas numpy openpyxl pywin32`

> `pywin32` is only needed for the Excel/COM tab (Windows). Make sure the **bitness** (32/64-bit) of Python matches Excel.

### 2) Run the app

`streamlit run m2py_streamlit_app.py`


* **Paste M** tab: paste your M, hit **Convert**.
* **Excel (Windows/COM)** tab: upload a workbook (`.xlsx`/`.xlsm`). The app:

  * lists queries,
  * shows the **dependency chain**,
  * converts to a runnable Python bundle,
  * prepends a `__cw` dict with any `Excel.CurrentWorkbook(){[Name="..."]}[Content]` tables it found.

Download the bundle with one click.

### 3) CLI (optional and largely untested, although all the logic is common)

```bash
python m2py_cli.py --in input.m --out output.py --name Result
```

---

## Files

* `m2py_streamlit_app.py` — Streamlit UI, COM reading, `__cw` preamble generation, dependency-aware bundling.
* `m2py_core.py` — core M→pandas conversion logic.
* `query_resolver.py` — finds query references and returns a dependency chain (topological order).
* `excel_com_extractor.py` — Windows/COM helpers:

  * `extract_queries_from_excel_via_com(path) -> {name: m_text}`
  * `extract_currentworkbook_tables_via_com(path, names) -> {Name: [row dicts]}`
* `pq_utils.py`, `project_utils.py`, `excel_com_extractor.py`, `query_resolver.py` — utilities used by core/app.

---

## Example

**M**:

```m
let
  Source = Table.FromRecords({[A=1, B="X"], [A=2, B="Y"], [A=3, B="X"]}),
  Filtered = Table.SelectRows(Source, each [B] = "X"),
  Sorted = Table.Sort(Filtered, {{"A", Order.Descending}})
in
  Sorted
```

**Python (excerpt)**:

```python
import pandas as pd
import numpy as np

Source = pd.DataFrame([{'A': 1, 'B': 'X'}, {'A': 2, 'B': 'Y'}, {'A': 3, 'B': 'X'}])
Filtered = Source[Source['B'] == 'X'].copy()
Sorted = Filtered.sort_values(by=['A'], ascending=[False]).reset_index(drop=True)
Result = Sorted
```

**Group example**:

```M
let
    Source = Table.FromRecords({[Cat="A", V=1],[Cat="A", V=3],[Cat="B", V=2],[Cat="B", V=8]}),
    Grouped = Table.Group(Source, {"Cat"}, {{"Avg", each List.Average([V]), type number}, {"Cnt", each List.Count([V]), Int64.Type}})
in
    Grouped
```

→

```python
gb = Source.groupby(['Cat'], dropna=False)
Grouped = gb.agg(Avg=('V', 'mean'), Cnt=('V', 'size')).reset_index()
```

---

## `Excel.CurrentWorkbook` and external Excel files

When the selected query (or its deps) uses `Excel.CurrentWorkbook(){[Name="..."]}[Content]`, the app prepends something like:

```python
import pandas as pd
__cw = {}
__cw['Table1'] = pd.DataFrame([{'Column1': 10}, {'Column1': 20}])  # materialized from your upload
```

…and **each converted step** guards against overwriting it:

```python
if '__cw' not in globals():
    __cw = {}  # filled by the app; maps Name -> DataFrame
```

The bundle also includes a **commented template** to load those tables from a normal `.xlsx` (no COM) using `openpyxl`. Users can just uncomment:

```python
# --- Optional: load CurrentWorkbook tables from an external Excel file ---
# EXCEL_PATH = r'C:\path\to\workbook.xlsx'
# from openpyxl import load_workbook
# from openpyxl.utils import range_boundaries
# wb = load_workbook(EXCEL_PATH, data_only=True)
#
# def _table_df(wb, table_name):
#     for ws in wb.worksheets:
#         for t in getattr(ws, '_tables', {}).values():
#             if t.name == table_name:
#                 min_col, min_row, max_col, max_row = range_boundaries(t.ref)
#                 data = list(ws.iter_rows(min_row=min_row, max_row=max_row,
#                                         min_col=min_col, max_col=max_col, values_only=True))
#                 if not data:
#                     return pd.DataFrame()
#                 return pd.DataFrame(data[1:], columns=data[0])  # first row = headers
#     return pd.DataFrame()
# __cw['Table1'] = _table_df(wb, 'Table1')  # uncomment and adapt
# # wb.close()
```

---

## Limitations / notes

* Not a full M parser. Complex custom functions, nested `each` lambdas, advanced type systems, and exotic connectors are out of scope for now.
* COM extraction is **Windows-only** and requires Excel + `pywin32`.
* `Excel.CurrentWorkbook` returns **empty DataFrames** if tables have no data rows (header-only tables have `DataBodyRange = None`).
* Error handling aims for “never crash; leave `# Unsupported:` comments instead.”

---

## Contributing

PRs with:

* new pattern handlers (with tiny test examples),
* bug repros (include minimal M),
* improved type inference,

…are very welcome.

---

## License

MIT (see `LICENSE`).
