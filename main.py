from pathlib import Path
import fitz

from datetime import datetime
import pandas as pd

excel_path = r"C:\Users\GreenhouseProduction\Downloads\US Customs Invoices Feb 10\Tester1 (5).xlsx"

for sheet in pd.ExcelFile(excel_path).sheet_names:
#read everything as strings so nothing becomes NaN or floats
    df_raw = pd.read_excel(excel_path, sheet_name=sheet, header=None, dtype=str, keep_default_na=False, engine="openpyxl") #change sheet_name to variable with loop later

    #shipping details grabbed from the excel sheet
    filename = df_raw.iat[1, 4]
    name = df_raw.iat[1, 2]
    invoice = df_raw.iat[1, 7]
    address = df_raw.iat[3, 2]
    city = df_raw.iat[4, 2]
    postalcode = df_raw.iat[5, 2]

    def data_entry_data(): #dictionary of shipping details to be entred in fields
        return {
    '13': name,
    '14': address,
    '15': city,
    '16': postalcode,
    '26': invoice
        }

    #point to the folder where your PDF lives
    document_dir = Path(r"C:\Users\GreenhouseProduction\Downloads\US Customs Invoices Feb 10")
    #file names
    source_file_name = "Blank Template copy - Copy.pdf"
    output_file_name =  filename + ".pdf"

    #build full paths
    source_file = document_dir / source_file_name
    output_file = document_dir / output_file_name

    data_entry_data = data_entry_data()

    ##filling out item details##
    doc = fitz.open(source_file)

    #collect items from excel sheet
    product_description = []
    quantity = []
    weight = []
    amount = []

    item_row = 7
    num_items = 0 #count number of items to be entered
    while df_raw.iat[item_row, 1] != '':
        product_description.append(str(df_raw.iat[item_row, 1]) + ' ' + str(df_raw.iat[item_row, 2]))
        item_row += 1

    num_items = len(product_description)

    for i in range(num_items):
        quantity.append(float(df_raw.iat[7+i, 6]))

    for i in range(num_items):
        weight.append(float(df_raw.iat[7+i, 7]))

    for i in range(num_items):
        amount.append(float(df_raw.iat[7+i, 9]))

    data_idx = 0 #number to track index in lists of data
    widget_counter = 0 #number to track widget index
    start_at = 48 #for the weight field

    for page_num, page in enumerate(doc):

        widgets = list(page.widgets())

        if page_num == 0:
                start_at = 48 #for the weight field on page one
        elif page_num == 1:
                start_at = 1 #for weight field on page two

        widget_counter = 0 #resets the widget counter for each page

        while widget_counter < len(widgets):

            if data_idx >= len(product_description):
                break

            if page_num == 0 and widget_counter < start_at: #fill out the shipping details first
                if widgets[widget_counter].field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                    widgets[widget_counter].field_value = data_entry_data.get(str(widget_counter), '')
                    widgets[widget_counter].update()

            if widget_counter >= start_at and (widget_counter - start_at) % 8 == 0: #fill out quantity, weight, amount
                
                if widget_counter >= 100:
                    break #go to next page after finishing the items on the first page

                widgets = list(page.widgets())

                if widgets[widget_counter].field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                    
                    widgets[widget_counter].field_value = str(weight[data_idx])
                    widgets[widget_counter].update()

                    widgets[widget_counter + 1].field_value = str(quantity[data_idx])
                    widgets[widget_counter + 1].update()

                    widgets[widget_counter + 2].field_value = str(amount[data_idx])
                    widgets[widget_counter + 2].update()

                    data_idx += 1
            
            widget_counter += 1
        
    #another loop through the pages to add the description of goods
    data_idx = 0 #reset data index to start of list of descriptions
    widget_counter = 0 #reset widget counter to start of page
    # for page_num, page in enumerate(doc):
    #     widgets = list(page.widgets())

    #     if page_num == 0:
    #             start_at = 48 #for the weight field on page one
    #     elif page_num == 1:
    #             start_at = 1 #for weight field on page two

    #     widget_counter = 0 #resets the widget counter for each page

    #     while widget_counter < len(widgets):
    #         print('fill in later')


    #updating other fields (date, total amount/quantity/weight)
    for page_num, page in enumerate(doc):

        widgets = list(page.widgets())
        # widget_counter = 0 #resets the widget counter for each page

        # while widget_counter < len(widgets):
        if page_num == 0:
            widgets[38].field_value = str(sum(amount))
            widgets[38].update() #total amount field (first instance)
            
            widgets[151].field_value = str(sum(weight))
            widgets[151].update() #total weight field

            widgets[152].field_value = str(sum(quantity))
            widgets[152].update() #total quantity field

            widgets[153].field_value = str(sum(amount))
            widgets[153].update() #total amount field

            widgets[155].field_value = datetime.now().strftime("%m/%d/%Y")
            widgets[155].update() #date field
            
            #widget_counter += 1

    doc.save(output_file)