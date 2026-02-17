import streamlit as st
import main
import io
import zipfile

st.title("Fill out U.S. Shipping Invoices")

uploaded_excel = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_excel and st.button("Generate Invoices"):
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer,"w") as zip_file:
        invoices = main.fill_out_US_shipping_invoices(uploaded_excel)

        for filename, pdf_bytes in invoices:
            zip_file.writestr(filename, pdf_bytes)

    zip_buffer.seek(0)

    st.download_button("Download All Invoices", data=zip_buffer, file_name="invoices.zip", mime="application/zip")

# if 'fill_out_invoices' not in st.session_state:
#     st.session_state.fill_out_invoices = ''

# def fill_out_invoices():
#     excel_file_path =  st.session_state.a
#     folder_path = st.session_state.b

#     main.fill_out_US_shipping_invoices(excel_file_path, folder_path)

#     st.session_state.fill_out_invoices = 'Finished Filling out Invoices'

# col1,col2 = st.columns(2)
# col1.title('Fill out U.S. Shipping Invoices')
# if st.session_state.fill_out_invoices != "":
#     col2.title('Finished Filling out Invoices')

# with st.form('invoice_information'):
#     st.text_input('Path to Excel workbook with shipping information:', key = 'a')
#     st.text_input('Path to Folder for Shipping Invoices (Blank template must be in here)', key = 'b')
#     st.form_submit_button('Submit', on_click=fill_out_invoices)