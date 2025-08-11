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
"""
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
