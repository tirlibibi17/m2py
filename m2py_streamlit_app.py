
import re
import tempfile
import os
import platform
import hashlib
import pathlib

import streamlit as st
from m2py_core import convert_m_to_python


# Import dependency resolver (deduplicated)
from query_resolver import find_query_refs, topo_order_queries, dependency_chain_for

# Excel COM extractor (Windows-only) — lazy import with fallback
try:
    from excel_com_extractor import extract_queries_from_excel_via_com  # noqa: F401
except Exception:  # pragma: no cover
    def extract_queries_from_excel_via_com(*args, **kwargs):  # type: ignore
        raise SystemExit("Excel COM extraction requires Windows + Excel + pywin32.")
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
