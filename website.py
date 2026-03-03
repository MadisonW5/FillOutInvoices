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

        if type(invoices) is bool and invoices == True:
            st.header("You forgot to delete an empty page in the Excel sheet. Edit the sheet and try again.")

        elif type(invoices) is bool and invoices != True:
            st.header("Make sure quantity, weight, and amount values are formatted as numbers in Excel file and ensure the formatting of the sheet hasn't changed. (refer to SOP)")

        else:
            for filename, pdf_bytes in invoices:
                zip_file.writestr(filename, pdf_bytes)

    if type(invoices) is not bool:
        zip_buffer.seek(0)

        st.header("Invoices Complete")

        st.download_button("Download Filled Invoices", data=zip_buffer, file_name="US Customs Invoices " + datetime.now().strftime("%m-%d-%Y") + ".zip", mime="application/zip")