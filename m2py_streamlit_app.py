import os
import platform
import tempfile
import re
import traceback
from pathlib import Path

import streamlit as st

from m2py_core import convert_m_to_python
from query_resolver import dependency_chain_for

# Try to import COM helpers. The app still runs without them (Paste M tab works).
try:
    from excel_com_extractor import (
        extract_queries_from_excel_via_com,
        extract_currentworkbook_tables_via_com,
    )  # type: ignore
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False


st.set_page_config(page_title="Power Query (M) → pandas converter", layout="wide")
st.title("Power Query (M) → pandas (Python) Converter")
st.caption(
    "Paste M code and convert to pandas, or pull queries from an Excel workbook via Windows COM. "
    "This is a pragmatic converter for prototyping—not a full M parser."
)

# ==============================
# Sidebar: curated examples
# ==============================
EXAMPLES = {
    "FromRecords (flat)": """let
    Source = Table.FromRecords({[A=1, B="x"], [A=2, B="y"], [A=3, B="z"]})
in
    Source
""",

    "Filter & Sort": """let
    Source = Table.FromRecords({[A=1, B="X"], [A=2, B="Y"], [A=3, B="X"]}),
    Filtered = Table.SelectRows(Source, each [B] = "X"),
    Sorted = Table.Sort(Filtered, {{"A", Order.Descending}})
in
    Sorted
""",

    "CSV + PromoteHeaders + Types": """let
    Source = Csv.Document(File.Contents("data.csv"), [Delimiter=";", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source),
    Types = Table.TransformColumnTypes(Promoted, {{"col1", type text}, {"col2", type number}})
in
    Types
""",

    "#table (numbers)": """let
    Source = #table({"A","B"}, {{1,2},{3,4}})
in
    Source
""",

    "Join (Inner, multi-key)": """let
    A = Table.FromRecords({[k1="a", k2=1, v1=10],[k1="b", k2=2, v1=20]}),
    B = Table.FromRecords({[k1="a", k2=1, v2=100],[k1="c", k2=3, v2=300]}),
    J = Table.Join(A, {"k1","k2"}, B, {"k1","k2"}, JoinKind.Inner)
in
    J
""",

    "Group: Sum": """let
    Source = Table.FromRecords({[A="X", Val=10],[A="X", Val=20],[A="Y", Val=5]}),
    Grouped = Table.Group(Source, {"A"}, {{"Sum", each List.Sum([Val]), type number}})
in
    Grouped
""",

    "Group: Avg + Count": """let
    Source = Table.FromRecords({[Cat="A", V=1],[Cat="A", V=3],[Cat="B", V=2],[Cat="B", V=8]}),
    Grouped = Table.Group(Source, {"Cat"}, {{"Avg", each List.Average([V]), type number}, {"Cnt", each List.Count([V]), Int64.Type}})
in
    Grouped
""",

    "Group: Min / Max / Median": """let
    Source = Table.FromRecords({[A="X", V=10],[A="X", V=20],[A="X", V=5],[A="Y", V=7]}),
    Grouped = Table.Group(Source, {"A"}, {{"Min", each List.Min([V]), type number}, {"Max", each List.Max([V]), type number}, {"Median", each List.Median([V]), type number}})
in
    Grouped
""",

    "Group: StdDev / Variance": """let
    Source = Table.FromRecords({[A="X", V=10],[A="X", V=20],[A="X", V=5],[A="Y", V=7],[A="Y", V=9]}),
    Grouped = Table.Group(Source, {"A"}, {{"Std", each List.StandardDeviation([V]), type number}, {"Var", each List.Variance([V]), type number}})
in
    Grouped
""",

    "Group: First / Last": """let
    Source = Table.FromRecords({[A="X", V=10],[A="X", V=20],[A="Y", V=5]}),
    Grouped = Table.Group(Source, {"A"}, {{"First", each List.First([V]), type any}, {"Last", each List.Last([V]), type any}})
in
    Grouped
""",

    "Group: Product": """let
    Source = Table.FromRecords({[A="X", V=2],[A="X", V=3],[A="Y", V=4]}),
    Grouped = Table.Group(Source, {"A"}, {{"Prod", each List.Product([V]), type number}})
in
    Grouped
""",
}

with st.sidebar:
    st.markdown("### Examples")
    if "m_input" not in st.session_state:
        st.session_state["m_input"] = ""
    ex_key = st.selectbox("Pick an example", list(EXAMPLES.keys()))
    st.code(EXAMPLES[ex_key], language="m")
    if st.button("Insert into editor"):
        st.session_state["m_input"] = EXAMPLES[ex_key]

# ==============================
# Tabs
# ==============================
tab1, tab2 = st.tabs(["Paste M", "Excel (Windows/COM)"])

# --------- Tab 1: Paste M ----------
with tab1:
    m_code = st.text_area("Paste your M code", height=320, key="m_input")
    if st.button("Convert", key="convert_m"):
        if (m_code or "").strip():
            try:
                py = convert_m_to_python(m_code, query_name="Result")
                st.markdown("**Python**")
                st.code(py, language="python")
                st.download_button("Download Python", py, file_name="converted.py")
            except Exception as e:
                st.error("Conversion error")
                st.code("".join(traceback.format_exception(type(e), e, e.__traceback__)), language="text")
        else:
            st.warning("Please paste some M code first.")

# --------- Tab 2: Excel (Windows/COM) ----------
with tab2:
    st.markdown("**Note:** Excel COM extraction requires Windows + Excel + pywin32.")
    if platform.system() != "Windows":
        st.warning("You're not on Windows; Excel COM will not work here.")
    if not _COM_AVAILABLE:
        st.info("excel_com_extractor helpers not available in this environment.")

    # File upload only (no manual path; autoload when uploaded)
    up = st.file_uploader("Upload an Excel workbook (.xlsx/.xlsm)", type=["xlsx", "xlsm"])

    # Session state for Excel workflow
    if "excel_queries" not in st.session_state:
        st.session_state["excel_queries"] = None
    if "excel_tmp_path" not in st.session_state:
        st.session_state["excel_tmp_path"] = None
    if "excel_sig" not in st.session_state:
        st.session_state["excel_sig"] = None
    if "excel_q_sel" not in st.session_state:
        st.session_state["excel_q_sel"] = None

    # When a new file is uploaded, persist to temp and auto-load queries
    if (
        up is not None
        and platform.system() == "Windows"
        and _COM_AVAILABLE
    ):
        sig = (up.name, up.size)
        if st.session_state["excel_sig"] != sig:
            # Save upload to a temp file
            try:
                fd, tmp = tempfile.mkstemp(suffix=f"_{up.name}")
                with os.fdopen(fd, "wb") as f:
                    f.write(up.read())
                st.session_state["excel_tmp_path"] = tmp
                st.session_state["excel_sig"] = sig

                # Auto-extract queries
                try:
                    raw = extract_queries_from_excel_via_com(tmp)  # type: ignore

                    # Normalize names to non-empty strings (prevents None keys)
                    fixed = {}
                    unnamed = 1
                    for k, v in (raw or {}).items():
                        name = (str(k).strip() if k is not None else "")
                        if not name:
                            while True:
                                cand = f"Query_{unnamed}"
                                unnamed += 1
                                if cand not in fixed:
                                    name = cand
                                    break
                        fixed[name] = v or ""

                    st.session_state["excel_queries"] = fixed
                    st.session_state["excel_q_sel"] = None
                    st.success(f"Loaded {len(fixed)} queries from '{up.name}'.")
                except Exception as e:
                    st.session_state["excel_queries"] = None
                    st.error("Excel COM error while reading queries:")
                    st.code("".join(traceback.format_exception(type(e), e, e.__traceback__)), language="text")
            except Exception as e:
                st.error("Failed to persist upload:")
                st.code("".join(traceback.format_exception(type(e), e, e.__traceback__)), language="text")

    queries = st.session_state.get("excel_queries")

    # If we have queries, allow selection and auto-convert on selection
    if isinstance(queries, dict) and queries:
        qnames = sorted([q for q in queries.keys() if isinstance(q, str) and q.strip()])

        # default selection: keep last valid choice, else first
        sel = st.session_state.get("excel_q_sel")
        if not isinstance(sel, str) or sel not in qnames:
            # clear stale value so `index=` is honored on first render
            st.session_state.pop("excel_q_sel", None)
            default_idx = 0
        else:
            default_idx = qnames.index(sel)

        # The widget will now own excel_q_sel; don't assign to it manually later
        q_sel = st.selectbox("Select a query", qnames, index=default_idx, key="excel_q_sel")

        try:
            # Build dependency chain and show it (ignore any weird/empty nodes)
            try:
                chain_raw = dependency_chain_for(q_sel, queries)
            except Exception:
                chain_raw = [q_sel]
            chain = [n for n in chain_raw if isinstance(n, str) and n in queries and n.strip()]
            if q_sel not in chain:
                chain.append(q_sel)
            st.caption("Dependency chain (deps → target): " + " → ".join(chain))

            # Find Excel.CurrentWorkbook names used anywhere in the chain
            names_needed = set()
            pat = re.compile(r'Excel\.CurrentWorkbook\(\)\{\s*\[Name="([^"]+)"\]\s*\}\[Content\]')
            for name in chain:
                m_text = (queries.get(name) or "")
                for nm in pat.findall(m_text):
                    names_needed.add(nm)

            # Try to materialize __cw with real DataFrames from the workbook
            cw_code = "import pandas as pd\n__cw = {}\n\n"
            tables_debug = {}
            if names_needed and st.session_state.get("excel_tmp_path") and _COM_AVAILABLE:
                try:
                    tables = extract_currentworkbook_tables_via_com(
                        st.session_state["excel_tmp_path"], sorted(names_needed)
                    ) or {}
                    # Build a readable __cw literal; emit ALL names (even if empty)
                    lines = ["import pandas as pd", "__cw = {}"]
                    for nm in sorted(names_needed):
                        rows = tables.get(nm) or []
                        tables_debug[nm] = len(rows)
                        lines.append(f"__cw['{nm}'] = pd.DataFrame({rows!r})")
                    cw_code = "\n".join(lines) + "\n\n"
                except Exception as e:
                    st.warning("Could not materialize __cw from workbook; falling back to empty dict.")
                    st.code("".join(traceback.format_exception(type(e), e, e.__traceback__)), language="text")

            # Optional: show what we found
            if tables_debug:
                st.caption(
                    "CurrentWorkbook items (rows): " +
                    ", ".join(f"{k}={v}" for k, v in sorted(tables_debug.items()))
                )

            # Convert all steps in dependency order and prepend __cw preamble
            blocks = []
            for name in chain:
                try:
                    m_txt = (queries.get(name) or "")
                    safe_name = name or "Result"
                    py_txt = convert_m_to_python(m_txt, query_name=safe_name)
                    blocks.append(f"# === {safe_name} ===\n{py_txt}\n")
                except Exception as e:
                    raise RuntimeError(f"convert_m_to_python failed for query '{name}'") from e

            py_all = cw_code + "\n".join(blocks)

            # Display side-by-side
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**M Code (selected)**")
                st.code(queries.get(q_sel) or "", language="m")
            with c2:
                st.markdown("**Python (bundle)**")
                st.code(py_all, language="python")

            # Bundle download
            safe = "".join(c if c.isalnum() or c in "._- " else "_" for c in (q_sel or "converted")).strip() or "converted"
            st.download_button("Download Python bundle", py_all, file_name=f"{safe}.py")

            # Individual per-step downloads
            for name, code in zip(chain, blocks):
                base = "".join(c if c.isalnum() or c in "._- " else "_" for c in (name or "step")).strip() or "step"
                st.download_button(f"Download {name or 'step'}.py", code, file_name=f"{base}.py", key=f"dl_{base}")

        except Exception as e:
            st.error("Conversion error")
            st.code("".join(traceback.format_exception(type(e), e, e.__traceback__)), language="text")

    # (Optional) You can add a "Clear upload" button if you want to reset session state.
