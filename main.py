def fill_out_US_shipping_invoices(uploaded_excel):
    from pathlib import Path
    import fitz
    from datetime import datetime
    import pandas as pd
    import io

    template_path = Path("Blank Template copy - final.pdf")
    invoices = []
    xls = pd.ExcelFile(uploaded_excel)
    
    for sheet in xls.sheet_names:
    #read everything as strings so nothing becomes NaN or floats
        df_raw = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=str, keep_default_na=False, engine="openpyxl")
        #shipping details grabbed from the excel sheet
        filename = df_raw.iat[1, 4]
        name = df_raw.iat[1, 2]
        invoice = df_raw.iat[1, 7]
        address = df_raw.iat[3, 2]
        city = df_raw.iat[4, 2]
        postalcode = df_raw.iat[5, 2]

        def shipping_details(): #dictionary of shipping details to be entred in fields
            return {
        '13': name.replace("Customer Name: ", ""),
        '14': address,
        '15': city,
        '16': postalcode,
        '26': invoice.replace("Ext. Ref: ", "")
            }

        #name of each filled out invoice form
        output_file_name =  filename.replace("Ref: ", "") + ".pdf"

        ##filling out item details##
        doc = fitz.open(template_path)

        #collect items from excel sheet
        product_description = []
        product_name = [] #the name without the SKU at the front
        tariff_number = []
        quantity = []
        weight = []
        amount = []

        item_row = 7
        num_items = 0 #count number of items to be entered
        while df_raw.iat[item_row, 1].strip() != '':
            product_description.append(str(df_raw.iat[item_row, 1]) + ' ' + str(df_raw.iat[item_row, 2]))
            product_name.append(str(df_raw.iat[item_row, 2]))
            item_row += 1

        num_items = len(product_description)

        for i in range(num_items):
            tariff_number.append(str(df_raw.iat[7+i, 5]))

        for i in range(num_items):
            quantity.append(float(df_raw.iat[7+i, 6]))

        for i in range(num_items):
            weight.append(float(df_raw.iat[7+i, 7]))

        for i in range(num_items):
            amount.append(float(df_raw.iat[7+i, 9]))

    #loop to add shipping details
        data_idx = 0 #number to track index in lists of data
        widget_counter = 0 #number to track widget index
        start_at = 48 #index for weight field on page one (end of shipping info is here)

        for page_num, page in enumerate(doc):

            widgets = list(page.widgets())

            widget_counter = 0 #resets the widget counter for each page

            while widget_counter < len(widgets):

                if page_num == 0 and widget_counter < start_at: #fill out the shipping details first
                    if widgets[widget_counter].field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                        widgets[widget_counter].field_value = shipping_details().get(str(widget_counter), '')
                        widgets[widget_counter].update()

                widget_counter += 1

    #loop to add weight, quantity, amount
        data_idx = 0 #number to track index in lists of data
        widget_counter = 0 #number to track widget index

        for page_num, page in enumerate(doc):

            widgets = list(page.widgets())

            if page_num == 0:
                    start_at = 48 #for the weight field on page one
                    end_of_page = 100
                    rows_per_page = 6 #number of items per page
            elif page_num == 1:
                    start_at = 1 #for weight field on page two
                    end_of_page = 155
                    rows_per_page = 25 #number of items per page

            widget_counter = 0 #resets the widget counter for each page

            for row in range(rows_per_page):
                widget_counter = start_at + row*8 #weight field is in every 8th field after the initial weight field

                if data_idx >= len(product_description):
                    break
                    
                if widget_counter + 2 >= end_of_page:
                    break #go to next page after finishing the items on the first page

                if widgets[widget_counter].field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                        
                    widgets[widget_counter].field_value = str(weight[data_idx])
                    widgets[widget_counter].update()

                    widgets[widget_counter + 1].field_value = str(quantity[data_idx])
                    widgets[widget_counter + 1].update()

                    widgets[widget_counter + 2].field_value = str(amount[data_idx])
                    widgets[widget_counter + 2].update()

                    data_idx += 1

        #another loop through the pages to add the description of goods
        data_idx = 0 #reset data index to start of list of descriptions
        widget_counter = 0 #reset widget counter to start of page

        for page_num, page in enumerate(doc):
            widgets = list(page.widgets())

            if page_num == 0:
                    start_at = 132 #for the description field on page one
                    end_at = 137
                    add_to_tariff = 6 #number of fields between description and tariff number fields

            elif page_num == 1:
                    start_at = 209 #for description field on page two
                    end_at = 227
                    add_to_tariff = 19 #number of fields between description and tariff number fields

            widget_counter = 0 #resets the widget counter for each page

            while widget_counter < len(widgets):
                
                if data_idx >= len(product_description):
                    break

                if widget_counter >= start_at and widget_counter <= end_at: #fill out description + tariff number
                    
                    if widget_counter > end_at:
                        break #go to next page after finishing the items on the first page

                    if widgets[widget_counter].field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                        
                        widgets[widget_counter].field_value = str(product_description[data_idx])
                        widgets[widget_counter].update()

                        widgets[widget_counter + add_to_tariff].field_value = str(tariff_number[data_idx])
                        widgets[widget_counter + add_to_tariff].update()

                    data_idx += 1

                widget_counter += 1

    ###
    #one more loop through the pages to add the second line of descriptions for goods (ex. "900 mL bottles, 6 bottles per box,  boxes")
        data_idx = 0 #reset data index to start of list of descriptions
        widget_counter = 0 #reset widget counter to start of page

        for page_num, page in enumerate(doc):
            widgets = list(page.widgets())

            if page_num == 0:
                    start_at = 145 #for the description field on page one
                    end_at = 150

            elif page_num == 1:
                    start_at = 247 #for description field on page two
                    end_at = 265

            widget_counter = 0 #resets the widget counter for each page

            while widget_counter < len(widgets):
                
                if data_idx >= len(product_description):
                    break

                if widget_counter >= start_at and widget_counter <= end_at: #fill out second line of description
                    
                    if widget_counter > end_at:
                        break #go to next page after finishing the items on the first page

                    if widgets[widget_counter].field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                        
                        if "6 x Gable Top" in str(product_name[data_idx]):
                            widgets[widget_counter].field_value = "900 mL bottles, 6 bottles per box," + " " + str(int(quantity[data_idx])) + " boxes"
                            widgets[widget_counter].update()

                        elif "2 x 6" in str(product_name[data_idx]) or "12 x" in str(product_name[data_idx]):
                            widgets[widget_counter].field_value = "60mL bottles, 12 bottles per box," + " " + str(int(quantity[data_idx])) + " boxes"
                            widgets[widget_counter].update()

                        elif "6 x 4" in str(product_name[data_idx]) or "24 x" in str(product_name[data_idx]) or "4 x 60 mL" in str(product_name[data_idx]):
                            widgets[widget_counter].field_value = "60mL bottles, 24 bottles per box," + " " + str(int(quantity[data_idx])) + " boxes"
                            widgets[widget_counter].update()

                        elif "6 x" in str(product_name[data_idx]) and "300 mL" in str(product_name[data_idx]):
                            widgets[widget_counter].field_value = "300 mL bottles, 6 bottles per box," + " " + str(int(quantity[data_idx])) + " boxes"
                            widgets[widget_counter].update()

                        elif "1.26 L" in str(product_name[data_idx]):
                            widgets[widget_counter].field_value = "1.26 L bag , 6 bags per box," + " " + str(int(quantity[data_idx])) + " boxes"
                            widgets[widget_counter].update()

                    data_idx += 1

                widget_counter += 1

        #updating other fields (date, total amount/quantity/weight)
        for page_num, page in enumerate(doc):

            widgets = list(page.widgets())
            # widget_counter = 0 #resets the widget counter for each page

            # while widget_counter < len(widgets):
            if page_num == 0:
                widgets[38].field_value = ""
                widgets[38].update() #total amount field (empty bc only if freight included)
                
                widgets[127].field_value = str(sum(weight))
                widgets[127].update() #total weight field

                widgets[128].field_value = str(sum(quantity))
                widgets[128].update() #total quantity field

                widgets[129].field_value = str(sum(amount))
                widgets[129].update() #total amount field

                widgets[131].field_value = datetime.now().strftime("%m/%d/%Y")
                widgets[131].update() #date field
                
                #widget_counter += 1

        pdf_bytes = doc.write()
        invoices.append((output_file_name, pdf_bytes))

    return invoices