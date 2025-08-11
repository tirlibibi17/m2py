import os
import platform
import tempfile
from pathlib import Path

import streamlit as st
from m2py_core import convert_m_to_python
from query_resolver import find_query_refs, topo_order_queries, dependency_chain_for

try:
    from excel_com_extractor import extract_queries_from_excel_via_com  # type: ignore
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False
    def extract_queries_from_excel_via_com(*args, **kwargs):  # type: ignore
        raise RuntimeError("Excel COM extraction requires Windows + Excel + pywin32.")

st.set_page_config(page_title="M to Python converter", layout="wide")
st.title("Power Query M -> pandas (Python) Converter")
st.caption(
    "Paste M code and convert to pandas, or read queries from an Excel workbook via Windows COM. "
    "This is a pragmatic converter for prototyping—not a full M parser."
)

# --- Examples in the sidebar -------------------------------------------------
EXAMPLES = {
    "Enter Data (Key/Value)": """let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45Wyk6tVDBU0lEqS8wpTQWyYnUgYkZwMSOl2FgA", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Key = _t, Value = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Key", type text}, {"Value", type text}})
in
    Source
""",

    "CSV + PromoteHeaders + Types": """let
    Source = Csv.Document(File.Contents("data.csv"), [Delimiter=";", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source),
    Types = Table.TransformColumnTypes(Promoted, {{"col1", type text}, {"col2", type number}})
in
    Types
""",

    "Filter & Sort": """let
    Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Amount", type number}, {"Region", type text}}),
    #"Filtered Rows" = Table.SelectRows(Source, each [Amount] > 1000 and [Region] <> "EMEA"),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Amount", Order.Descending}})
in
    Source
""",

    "Group & Aggregate": """let
    Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Region", type text}, {"Amount", type number}}),
    #"Grouped Rows" = Table.Group(Source, {"Region"}, {{"Total", each List.Sum([Amount]), type number}, {"Count", each Table.RowCount(_), Int64.Type}})
in
    Source
""",

    "Unpivot": """let
    Source = Excel.CurrentWorkbook(){[Name="PivotData"]}[Content],
    #"Unpivoted" = Table.UnpivotOtherColumns(Source, {"Key"}, "Attribute", "Value")
in
    Source
""",

    "Join (Inner, single key)": """let
    A = Excel.CurrentWorkbook(){[Name="A"]}[Content],
    B = Excel.CurrentWorkbook(){[Name="B"]}[Content],
    #"Changed Type A" = Table.TransformColumnTypes(A,{{"k", type text}}),
    #"Changed Type B" = Table.TransformColumnTypes(B,{{"k", type text}}),
    Joined = Table.Join(A, "k", B, "k", "Bcols", JoinKind.Inner)
in
    Joined
""",

    "Join (Inner, multi-key)": """let
    A = Excel.CurrentWorkbook(){[Name="A"]}[Content],
    B = Excel.CurrentWorkbook(){[Name="B"]}[Content],
    #"Changed Type A" = Table.TransformColumnTypes(A,{{"k1", type text}, {"k2", type number}}),
    #"Changed Type B" = Table.TransformColumnTypes(B,{{"k1", type text}, {"k2", type number}}),
    Joined = Table.Join(A, {"k1","k2"}, B, {"k1","k2"}, "Bcols", JoinKind.Inner)
in
    Joined
"""
}

with st.sidebar:
    st.markdown("### Examples (starter snippets)")
    ex = st.selectbox("Pick an example", list(EXAMPLES.keys()))
    st.code(EXAMPLES[ex], language="m")
    if st.button("Insert into editor"):
        st.session_state["m_input"] = EXAMPLES[ex]

tab1, tab2 = st.tabs(["Paste M", "Excel (Windows/COM)"])

with tab1:
    m_code = st.text_area("Paste your M code", height=320, key="m_input")
    convert = st.button("Convert", key="convert_m")
    if convert:
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

    up = st.file_uploader("Upload an Excel workbook (.xlsx/.xlsm)", type=["xlsx","xlsm"])
    path_text = st.text_input("...or provide a full path to a local Excel workbook")

    queries = None
    tmp_path = None

    if up is not None:
        try:
            fd, tmp = tempfile.mkstemp(suffix=f"_{up.name}")
            with os.fdopen(fd, "wb") as f:
                f.write(up.read())
            tmp_path = tmp
            st.caption(f"Saved uploaded workbook to: {tmp_path}")
        except Exception as e:
            st.error(f"Failed to persist upload: {e}")

    target_path = path_text.strip() or tmp_path or ""

    if st.button("Read queries from workbook", disabled=not target_path):
        try:
            queries = extract_queries_from_excel_via_com(target_path)
        except Exception as e:
            st.error(f"Excel COM error: {e}")

    if "excel_queries" not in st.session_state:
        st.session_state["excel_queries"] = None
    if queries is not None:
        st.session_state["excel_queries"] = queries
    queries = st.session_state.get("excel_queries")

    if isinstance(queries, dict) and queries:
        st.success(f"Loaded {len(queries)} queries.")
        qnames = sorted(queries.keys())
        q_sel = st.selectbox("Select a query", qnames)

        chain = dependency_chain_for(q_sel, queries)
        st.caption("Dependency chain (deps -> target): " + " -> ".join(chain))

        if st.button("Convert selected query (with deps)"):
            try:
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
                safe = "".join(c if c.isalnum() or c in '._- ' else '_' for c in q_sel).strip() or "converted"
                st.download_button("Download Python bundle", py_all, file_name=f"{safe}.py")
            except Exception as e:
                st.error(f"Conversion error: {e}")

    if tmp_path and Path(tmp_path).exists():
        try:
            os.remove(tmp_path)
        except Exception:
            pass
