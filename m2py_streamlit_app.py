
import re
import tempfile
import os
import platform
import hashlib
import pathlib

import streamlit as st
from m2py_core import convert_m_to_python

# --------- inline dependency resolver (same as CLI) ---------

def _strip_line_comments(m_code: str) -> str:
    cleaned_lines = []
    for line in (m_code or "").splitlines():
        s = line
        out, i, in_str, quote = [], 0, False, ""
        while i < len(s):
            ch = s[i]
            if not in_str and ch in ('"', "'"):
                in_str, quote = True, ch
                out.append(ch); i += 1; continue
            if in_str:
                out.append(ch)
                if ch == quote:
                    in_str, quote = False, ""
                i += 1; continue
            if ch == "/" and i + 1 < len(s) and s[i + 1] == "/":
                break
            out.append(ch); i += 1
        cleaned_lines.append("".join(out))
    return "\n".join(cleaned_lines)

REF_QUOTED = re.compile(r'#"(.*?)"')
IDENT = re.compile(r'\b([A-Za-z_][A-Za-z0-9_]*)\b')
M_KEYWORDS = {"let","in","each","and","or","not","true","false","null","as","if","then","else","error","try","otherwise"}

def find_query_refs(m_code: str, known_names: set[str]) -> set[str]:
    src = _strip_line_comments(m_code or "")
    refs = set(re.findall(r'#"(.*?)"', src))
    cand = set(IDENT.findall(src))
    refs |= {n for n in cand if n in known_names and n.lower() not in M_KEYWORDS}
    return refs

def topo_order_queries(queries: dict[str,str]) -> list[str]:
    known = set(queries.keys())
    graph = {name: {r for r in find_query_refs(m, known) if r in known and r != name}
             for name, m in queries.items()}
    indeg = {n: 0 for n in queries}
    for n, deps in graph.items():
        for _ in deps: indeg[n] += 1
    from collections import deque
    q = deque([n for n,d in indeg.items() if d==0])
    order = []
    while q:
        u = q.popleft()
        order.append(u)
        for v,deps in graph.items():
            if u in deps:
                indeg[v] -= 1
                if indeg[v]==0 and v not in order and v not in q:
                    q.append(v)
    remaining = [n for n in queries if n not in order]
    return order + remaining

def dependency_chain_for(target: str, queries: dict[str,str]) -> list[str]:
    order = topo_order_queries(queries)
    known = set(queries.keys())
    from collections import defaultdict
    rev = defaultdict(set)
    for n,m in queries.items():
        for r in find_query_refs(m, known):
            if r in known and r != n: rev[n].add(r)
    seen = set(); stack=[target]
    while stack:
        cur = stack.pop()
        if cur in seen: continue
        seen.add(cur); stack.extend(rev[cur])
    return [n for n in order if n in seen]

# --------- inline Excel COM extractor ---------

def extract_queries_from_excel_via_com(path_xlsx: str) -> dict[str,str]:
    try:
        import gc, pythoncom
        import win32com.client as win32
        from contextlib import contextmanager
    except Exception:
        raise RuntimeError("Requires Windows + Excel + 'pywin32'. Install: pip install pywin32")

    @contextmanager
    def _com_apartment():
        pythoncom.CoInitialize()
        try:
            yield
        finally:
            pythoncom.CoUninitialize()

    path_xlsx = str(pathlib.Path(path_xlsx).resolve())

    with _com_apartment():
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = None
        try:
            wb = excel.Workbooks.Open(path_xlsx, ReadOnly=True, UpdateLinks=0)
            result = {}
            for q in wb.Queries:
                result[str(q.Name)] = str(q.Formula)
            return result
        finally:
            try:
                if wb is not None: wb.Close(False)
            except Exception: pass
            try:
                excel.Quit()
            except Exception: pass
            del wb; del excel; gc.collect()

# --------- UI ---------

st.set_page_config(page_title="M → pandas Converter", layout="wide")
st.title("Power Query (M) → pandas Converter")

tab1, tab2 = st.tabs(["🔤 Paste M", "📗 Excel (Windows/COM)"])

# --- Paste M ---
with tab1:
    m_code = st.text_area("Paste your M code", height=300)
    if st.button("Convert", key="convert_m"):
        if m_code.strip():
            try:
                py = convert_m_to_python(m_code, query_name="Result")  # alias appended
                st.markdown("**Python**")
                st.code(py, language="python")
                st.download_button("Download Python", py, file_name="converted.py")
            except Exception as e:
                st.error(f"Conversion error: {e}")
        else:
            st.warning("Please paste some M code first.")

# --- Excel via COM (Windows only) ---
with tab2:
    st.write("Upload an Excel workbook (.xlsx/.xlsm) that contains Power Query queries.")
    if platform.system() != "Windows":
        st.info("This feature requires Windows + Excel (COM automation).")

    uploaded = st.file_uploader("Upload Excel file", type=["xlsx", "xlsm"])

    if uploaded is not None:
        buf = uploaded.getbuffer()
        sig = hashlib.sha1(buf).hexdigest()

        if st.session_state.get("excel_sig") != sig:
            with tempfile.TemporaryDirectory() as td:
                path = os.path.join(td, uploaded.name)
                with open(path, "wb") as f:
                    f.write(buf)
                try:
                    queries = extract_queries_from_excel_via_com(path)
                except Exception as e:
                    st.error(f"Failed to extract queries via Excel COM: {e}")
                    st.stop()
            st.session_state["excel_sig"] = sig
            st.session_state["excel_queries"] = queries

        queries = st.session_state.get("excel_queries", {})

        if not queries:
            st.warning("No queries found in this workbook.")
        else:
            qname = st.selectbox("Select a query", sorted(queries.keys()))
            include_deps = st.toggle("Include dependencies", value=True)

            if include_deps:
                chain = dependency_chain_for(qname, queries)  # deps first, then target
                blocks: list[str] = []
                for name in chain:
                    m_text = queries[name]
                    py_text = convert_m_to_python(m_text, query_name=name)  # alias appended
                    blocks.append(f"# === {name} ===\n{py_text}\n")
                py = "\n".join(blocks)
            else:
                m_text = queries[qname]
                py = convert_m_to_python(m_text, query_name=qname)

            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**M Code**")
                st.code(queries[qname], language="m")
            with c2:
                st.markdown("**Python**")
                st.code(py, language="python")
            st.download_button("Download Python", py, file_name=f"{qname}.py")
