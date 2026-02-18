import streamlit as st
import main
import io
import zipfile
from datetime import datetime

st.title("Fill out U.S. Shipping Invoices")

uploaded_excel = st.file_uploader("Upload Excel File with Invoice Information", type=["xlsx"])

if uploaded_excel and st.button("Fill out Invoices"):
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer,"w") as zip_file:
        invoices = main.fill_out_US_shipping_invoices(uploaded_excel)

        for filename, pdf_bytes in invoices:
            zip_file.writestr(filename, pdf_bytes)

    zip_buffer.seek(0)

    st.header("Invoices Complete")

    st.download_button("Download Filled Invoices", data=zip_buffer, file_name="invoices " + datetime.now().strftime("%m/%d/%Y") + ".zip", mime="application/zip")