import os
import platform
import tempfile
from pathlib import Path

import streamlit as st

# If you prefer hot-reload during dev, uncomment these 3 lines:
# import importlib, m2py_core
# importlib.reload(m2py_core)
# convert_m_to_python = m2py_core.convert_m_to_python

from m2py_core import convert_m_to_python
from query_resolver import find_query_refs, topo_order_queries, dependency_chain_for

try:
    from excel_com_extractor import extract_queries_from_excel_via_com  # type: ignore
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False

st.set_page_config(page_title="Power Query M → pandas converter", layout="wide")
st.title("Power Query (M) → pandas (Python) Converter")
st.caption(
    "Paste M code and convert to pandas, or read queries from an Excel workbook via Windows COM. "
    "This is a pragmatic converter for prototyping—not a full M parser."
)

# --- Examples in the sidebar -------------------------------------------------
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

# --- Main content: two tabs --------------------------------------------------
tab1, tab2 = st.tabs(["Paste M", "Excel (Windows/COM)"])

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
                st.error(f"Conversion error: {e}")
        else:
            st.warning("Please paste some M code first.")

with tab2:
    st.markdown("**Note:** Excel COM extraction requires Windows + Excel + pywin32.")
    if platform.system() != "Windows":
        st.warning("You're not on Windows; Excel COM will not work here.")
    if not _COM_AVAILABLE:
        st.info("excel_com_extractor not available in this environment.")

    # File upload only (no manual path)
    up = st.file_uploader("Upload an Excel workbook (.xlsx/.xlsm)", type=["xlsx", "xlsm"])

    # Auto-load queries when a file is uploaded
    if "excel_queries" not in st.session_state:
        st.session_state["excel_queries"] = None
    if "excel_tmp_path" not in st.session_state:
        st.session_state["excel_tmp_path"] = None
    if "excel_sig" not in st.session_state:
        st.session_state["excel_sig"] = None

    if up is not None and platform.system() == "Windows" and _COM_AVAILABLE:
        sig = (up.name, up.size)
        if st.session_state["excel_sig"] != sig:
            # Persist upload to a temp file
            try:
                fd, tmp = tempfile.mkstemp(suffix=f"_{up.name}")
                with os.fdopen(fd, "wb") as f:
                    f.write(up.read())
                st.session_state["excel_tmp_path"] = tmp
                st.session_state["excel_sig"] = sig
                # Extract queries immediately (autoload)
                try:
                    queries = extract_queries_from_excel_via_com(tmp)  # type: ignore
                    st.session_state["excel_queries"] = queries
                    st.success(f"Loaded {len(queries)} queries from '{up.name}'.")
                except Exception as e:
                    st.session_state["excel_queries"] = None
                    st.error(f"Excel COM error while reading queries: {e}")
            except Exception as e:
                st.error(f"Failed to persist upload: {e}")

    queries = st.session_state.get("excel_queries")

    # Auto-convert as soon as a query is selected
    if isinstance(queries, dict) and queries:
        qnames = sorted(queries.keys())
        # remember last selected
        default_idx = 0
        if "excel_q_sel" in st.session_state:
            try:
                default_idx = qnames.index(st.session_state["excel_q_sel"])
            except Exception:
                default_idx = 0
        q_sel = st.selectbox("Select a query", qnames, index=default_idx, key="excel_q_sel")

        # Build dependency chain and convert automatically
        try:
            chain = dependency_chain_for(q_sel, queries)
            st.caption("Dependency chain (deps → target): " + " → ".join(chain))

            blocks = []
            for name in chain:
                m_txt = queries[name]
                py_txt = convert_m_to_python(m_txt, query_name=name)
                blocks.append(f"# === {name} ===\n{py_txt}\n")

            py_all = "\n".join(blocks)

            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**M Code (selected)**")
                st.code(queries[q_sel], language="m")
            with c2:
                st.markdown("**Python (bundle)**")
                st.code(py_all, language="python")

            # Bundle download
            safe = "".join(c if c.isalnum() or c in "._- " else "_" for c in q_sel).strip() or "converted"
            st.download_button("Download Python bundle", py_all, file_name=f"{safe}.py")

            # Individual per-step downloads
            for name, code in zip(chain, blocks):
                base = "".join(c if c.isalnum() or c in "._- " else "_" for c in name).strip() or "step"
                st.download_button(f"Download {name}.py", code, file_name=f"{base}.py", key=f"dl_{name}")

        except Exception as e:
            st.error(f"Conversion error: {e}")

    # Optional: cleanup temp file when a new file replaces it (we keep current one for session lifetime)
    # You can also add a small "Clear" button if desired.
