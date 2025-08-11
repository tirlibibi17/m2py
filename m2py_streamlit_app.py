import streamlit as st
import pandas as pd
import numpy as np
import re

from m2py_core import convert_m_to_python

st.set_page_config(page_title="Power Query (M) to Pandas Converter", layout="wide")

# -------------------------
# Sample M code examples
# -------------------------
EXAMPLES = {
    "1. Table.FromRecords (flat)": """let
    Source = Table.FromRecords({
        [A=1, B="X"],
        [A=2, B="Y"]
    })
in
    Source""",

    "2. Table.FromRecords + ExpandRecordColumn": """let
    Source = Table.FromRecords({
        [A=1, Rec=[x=10, y=20]],
        [A=2, Rec=[x=30, y=40]]
    }),
    Expanded = Table.ExpandRecordColumn(Source, "Rec", {"x", "y"}, {"X", "Y"})
in
    Expanded""",

    "3. #table with numbers": """let
    Source = #table({"A", "B"}, {{1,2},{3,4}})
in
    Source""",

    "4. #table with text and numbers": """let
    Source = #table({"Key", "Value"}, {{"k1", 10}, {"k2", 20}})
in
    Source""",

    "5. CSV + PromoteHeaders + Types": """let
    Source = Csv.Document(File.Contents("data.csv"), [Delimiter=";", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Types = Table.TransformColumnTypes(Promoted, {{"col1", type text}, {"col2", Int64.Type}})
in
    Types""",

    "6. Filter and Sort": """let
    Source = Table.FromRecords({[A=1, B="X"], [A=2, B="Y"], [A=3, B="X"]}),
    Filtered = Table.SelectRows(Source, each [B] = "X"),
    Sorted = Table.Sort(Filtered, {{"A", Order.Descending}})
in
    Sorted""",

    "7. Group and Aggregate": """let
    Source = Table.FromRecords({[A="X", Val=10],[A="X", Val=20],[A="Y", Val=5]}),
    Grouped = Table.Group(Source, {"A"}, {{"Sum", each List.Sum([Val]), type number}})
in
    Grouped""",

    "8. Unpivot": """let
    Source = Table.FromRecords({[A=1, B=2, C=3], [A=4, B=5, C=6]}),
    Unpivoted = Table.Unpivot(Source, {"B", "C"}, "Attribute", "Value")
in
    Unpivoted""",

    "9. Join (single key)": """let
    Table1 = Table.FromRecords({[K=1, V1="A"], [K=2, V1="B"]}),
    Table2 = Table.FromRecords({[K=1, V2="X"], [K=3, V2="Y"]}),
    Joined = Table.Join(Table1, "K", Table2, "K", JoinKind.Inner)
in
    Joined""",

    "10. Join (multi-key)": """let
    Table1 = Table.FromRecords({[K1=1, K2="A", V1=100], [K1=2, K2="B", V1=200]}),
    Table2 = Table.FromRecords({[K1=1, K2="A", V2=300], [K1=3, K2="C", V2=400]}),
    Joined = Table.Join(Table1, {"K1", "K2"}, Table2, {"K1", "K2"}, JoinKind.Inner)
in
    Joined"""
}

# -------------------------
# Streamlit App
# -------------------------
st.title("Power Query (M) → Pandas Converter")

st.markdown("""
This tool converts Power Query M code into Python Pandas code.
""")

example_choice = st.selectbox("Select an example:", [""] + list(EXAMPLES.keys()))

if example_choice:
    m_code_input = EXAMPLES[example_choice]
else:
    m_code_input = ""

m_code = st.text_area("Paste your M code here:", m_code_input, height=300)

if st.button("Convert to Python"):
    if m_code.strip():
        py_code = convert_m_to_python(m_code)
        st.code(py_code, language="python")
    else:
        st.warning("Please paste some M code to convert.")

st.markdown("---")
st.caption("M → Pandas converter demo")
