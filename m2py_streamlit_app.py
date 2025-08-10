# m2py_streamlit_app.py
import streamlit as st
import tempfile, os, platform, hashlib
from m2py_core import convert_m_to_python

st.set_page_config(page_title="M → pandas Converter", layout="wide")
st.title("Power Query (M) → pandas Converter")

tab1, tab2 = st.tabs(["🔤 Paste M", "📗 Excel (Windows/COM)"])

# --- Paste M ---
with tab1:
    m_code = st.text_area("Paste your M code", height=300)
    if st.button("Convert", key="convert_m"):
        if m_code.strip():
            try:
                py = convert_m_to_python(m_code)
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
        # Compute a stable signature for the upload
        buf = uploaded.getbuffer()
        sig = hashlib.sha1(buf).hexdigest()

        if st.session_state.get("excel_sig") != sig:
            # New file => extract once via COM, store in session_state
            try:
                from excel_com_extractor import extract_queries_from_excel_via_com
            except ImportError:
                st.error("Missing dependency: 'pywin32'. Install with: pip install pywin32 (and Excel must be installed).")
                st.stop()

            with tempfile.TemporaryDirectory() as td:
                path = os.path.join(td, uploaded.name)
                with open(path, "wb") as f:
                    f.write(buf)  # save to disk for Excel COM

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
            m_text = queries[qname]

            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**M Code**")
                st.code(m_text, language="m")
            with c2:
                try:
                    py = convert_m_to_python(m_text)
                    st.markdown("**Python**")
                    st.code(py, language="python")
                    st.download_button("Download Python", py, file_name=f"{qname}.py")
                except Exception as e:
                    st.error(f"Conversion error: {e}")