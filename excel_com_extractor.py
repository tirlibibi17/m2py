# excel_com_extractor.py
# Helpers to pull Power Query M and CurrentWorkbook data via Excel COM (Windows).
# Requires: pywin32 (win32com + pythoncom), and Excel (bitness must match Python).

from __future__ import annotations
from typing import Dict, List, Iterable, Optional
import os

# pywin32 imports are only available on Windows
try:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
    from win32com.client import DispatchEx  # type: ignore
    from pywintypes import com_error  # type: ignore
except Exception:  # pragma: no cover
    pythoncom = None
    win32com = None
    DispatchEx = None
    com_error = Exception  # fallback so annotations still work


def _ensure_windows():
    import platform
    if platform.system() != "Windows":
        raise RuntimeError("Excel COM extraction requires Windows.")


def _open_workbook(path: str):
    """
    Open an Excel workbook read-only via COM and return (excel, workbook).
    Caller MUST Close/Quit in a finally block.
    """
    if not path or not os.path.exists(path):
        raise FileNotFoundError(f"Workbook not found: {path}")

    _ensure_windows()
    if win32com is None or DispatchEx is None:
        raise RuntimeError("pywin32 (win32com) is not available.")

    excel = DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(path, ReadOnly=True)
    return excel, wb


def _to_2d_list(val):
    """Normalize Excel Range.Value to a 2D list."""
    if val is None:
        return []
    if isinstance(val, tuple):
        if len(val) and isinstance(val[0], tuple):
            return [list(r) for r in val]
        return [list(val)]
    return [[val]]


# --------------------------
# Queries (Power Query M)
# --------------------------
def extract_queries_from_excel_via_com(workbook_path: str) -> Dict[str, str]:
    """
    Return a dict {QueryName: M_code} using Workbook.Queries (Excel 2016+).
    Names are normalized to non-empty strings; formulas default to "" if missing.
    COM is initialized per-call to work in Streamlit's worker threads.
    """
    # --- COM init ---
    if pythoncom is None:
        _ensure_windows()
        raise RuntimeError("pythoncom is not available; install pywin32.")
    # STA is fine for Excel automation
    pythoncom.CoInitialize()

    excel = wb = None
    out: Dict[str, str] = {}
    try:
        excel, wb = _open_workbook(workbook_path)

        # Preferred API: Workbook.Queries
        try:
            queries = wb.Queries
            count = int(getattr(queries, "Count", 0))
        except com_error:
            queries = None
            count = 0

        if queries and count > 0:
            for i in range(1, count + 1):  # 1-based COM indexing
                try:
                    q = queries.Item(i)
                    name = str(getattr(q, "Name", "") or "").strip()
                    if not name:
                        name = f"Query_{i}"
                    formula = str(getattr(q, "Formula", "") or "")
                    out[name] = formula
                except com_error:
                    # Skip problematic entries, keep going
                    continue
        else:
            # Fallback (older Excel): try connections (limited; may not expose M)
            try:
                conns = wb.Connections
                ccount = int(getattr(conns, "Count", 0))
                for i in range(1, ccount + 1):
                    try:
                        c = conns.Item(i)
                        name = str(getattr(c, "Name", "") or "").strip() or f"Connection_{i}"
                        out[name] = ""
                    except com_error:
                        continue
            except com_error:
                pass

        return out

    finally:
        # Cleanly close workbook and Excel
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        # --- COM uninit ---
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


# --------------------------------------
# CurrentWorkbook tables / named ranges
# --------------------------------------
def extract_currentworkbook_tables_via_com(
    workbook_path: str,
    names: Optional[Iterable[str]] = None
) -> Dict[str, List[dict]]:
    """
    Return {Name -> list-of-dicts} for Excel CurrentWorkbook items referenced by Power Query:
      - Excel Tables (ListObjects): column names from ListColumns.Name
      - Named Ranges: first row is headers, remaining rows are data
    If `names` is provided, only those names are fetched (missing names return empty lists).
    COM is initialized per-call to work in Streamlit's worker threads.
    """
    # --- COM init ---
    if pythoncom is None:
        _ensure_windows()
        raise RuntimeError("pythoncom is not available; install pywin32.")
    pythoncom.CoInitialize()

    excel = wb = None
    want = set(n.strip() for n in names) if names else None
    out: Dict[str, List[dict]] = {}

    try:
        excel, wb = _open_workbook(workbook_path)

        # -------- Tables (ListObjects) --------
        for ws in wb.Worksheets:
            try:
                list_objects = ws.ListObjects
                count = int(getattr(list_objects, "Count", 0))
            except com_error:
                count = 0

            for i in range(1, count + 1):
                try:
                    lo = list_objects.Item(i)
                    nm = str(getattr(lo, "Name", "") or "").strip()
                    if not nm:
                        continue
                    if want and nm not in want:
                        continue

                    # Headers: use ListColumns to be robust
                    headers = [str(col.Name) for col in lo.ListColumns]
                    rows: List[dict] = []
                    dbr = getattr(lo, "DataBodyRange", None)
                    if dbr is not None:
                        data = _to_2d_list(dbr.Value)
                        for r in data:
                            row = {}
                            for j, h in enumerate(headers):
                                row[h] = r[j] if j < len(r) else None
                            rows.append(row)

                    out[nm] = rows
                except com_error:
                    continue

        # -------- Named ranges (Workbook.Names) --------
        try:
            names_coll = wb.Names
            ncount = int(getattr(names_coll, "Count", 0))
        except com_error:
            names_coll = None
            ncount = 0

        for i in range(1, ncount + 1):
            try:
                nm_obj = names_coll.Item(i)
                nm = str(getattr(nm_obj, "Name", "") or "").strip()
                if not nm:
                    continue
                if want and nm not in want:
                    continue

                try:
                    rng = nm_obj.RefersToRange  # may raise if not a simple range
                except com_error:
                    rng = None
                if rng is None:
                    continue

                values = _to_2d_list(rng.Value)
                if not values:
                    out[nm] = []
                    continue

                headers = [str(x) if x is not None else "" for x in values[0]]
                data_rows = values[1:] if len(values) > 1 else []
                rows: List[dict] = []
                for r in data_rows:
                    row = {}
                    for j, h in enumerate(headers):
                        row[h] = r[j] if j < len(r) else None
                    rows.append(row)
                out[nm] = rows
            except com_error:
                continue

        # Ensure all requested names exist in output
        if want:
            for nm in want:
                out.setdefault(nm, [])

        return out

    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        # --- COM uninit ---
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
