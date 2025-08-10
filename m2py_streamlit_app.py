import streamlit as st
from m2py_core import convert_m_to_python
from pq_utils import extract_m_code_from_pq
from project_utils import save_project_zip, load_project_zip
import zipfile
import tempfile
import os

st.set_page_config(page_title="M to Python Converter", layout="wide")
st.title("Power Query (M) to Python Converter 🐼")

tab1, tab2 = st.tabs(["Paste M Code", "Upload .pq File"])

with tab1:
    m_code = st.text_area("M Code", height=300)
    if st.button("Convert"):
        try:
            python_code = convert_m_to_python(m_code)
            st.code(python_code, language="python")
            st.download_button("Download Python", python_code, file_name="converted.py")
        except Exception as e:
            st.error(str(e))

with tab2:
    uploaded_file = st.file_uploader("Upload a .pq file", type=["pq"])
    if uploaded_file:
        extracted = extract_m_code_from_pq(uploaded_file)
        for name, code in extracted:
            try:
                py = convert_m_to_python(code)
                st.subheader(name)
                st.code(py, language="python")
            except Exception as e:
                st.error(f"{name}: {e}")