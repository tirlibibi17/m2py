# excel_com_extractor.py
# Windows only. Requires: pip install pywin32  (and Excel installed)

import os
import gc
import pythoncom
import win32com.client as win32
from contextlib import contextmanager

@contextmanager
def _com_apartment():
    # Initialize COM for the current thread
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()

def extract_queries_from_excel_via_com(path_xlsx: str) -> dict:
    """
    Return {query_name: m_text} by reading Workbook.Queries via Excel COM.
    Safe to call from Streamlit re-runs (threaded).
    """
    path_xlsx = os.path.abspath(path_xlsx)
    with _com_apartment():
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = None
        try:
            # Avoid update links prompts
            wb = excel.Workbooks.Open(path_xlsx, ReadOnly=True, UpdateLinks=0)
            result = {}
            # Workbook.Queries is a COM collection
            for q in wb.Queries:
                # q.Name, q.Formula contain the M query metadata
                result[str(q.Name)] = str(q.Formula)
            return result
        finally:
            try:
                if wb is not None:
                    wb.Close(False)
            except Exception:
                pass
            try:
                excel.Quit()
            except Exception:
                pass
            # Release COM refs promptly
            del wb
            del excel
            gc.collect()
