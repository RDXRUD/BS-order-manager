import os
from openpyxl import load_workbook
import win32com.client
import pythoncom
import base64
import streamlit as st
from openpyxl.styles import Border, Side
from datetime import date
import pandas as pd
from datetime import datetime


def modify_excel_and_convert_to_pdf(input_file):
    workbook = load_workbook(input_file)
    
    sheet_names = workbook.sheetnames 
    # st.write(f"Sheet names: {sheet_names}")

    if len(sheet_names) > 1:
        for sheet_name in sheet_names[1:]:
            del workbook[sheet_name]

    sheet_names = workbook.sheetnames 
    # st.write(f"Remaining sheet names: {sheet_names}")

    sheet = workbook[sheet_names[0]]

    find_qty = False
    qty_column = None
    serial_count = 0
    blank_line = 0

    no_border = Border(left=Side(border_style=None),
                   right=Side(border_style=None),
                   top=Side(border_style=None),
                   bottom=Side(border_style=None))
    # Find 'Qty.' column
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None and isinstance(cell.value, str) and cell.value.strip().lower() == "qty.":
                find_qty = True
                qty_column = cell.column 
                # st.write(f"'qty' found at Row: {cell.row}, Column: {cell.column}")
                break
        if find_qty:
            break

    if qty_column:
        rows_to_delete = []

        # Modify rows based on conditions
        for row in sheet.iter_rows(min_row=cell.row + 1):
            qty_cell = row[qty_column - 1] 
            
            if row[1].value not in (None, '') or row[0].value not in (None, ''): 
                if not isinstance(qty_cell.value, int):
                    if not isinstance(row[0].value, str):
                        rows_to_delete.append(row[0].row) 
                elif qty_cell.value==0:
                    if not isinstance(row[0].value, str):
                        rows_to_delete.append(row[0].row) 
                else:
                    if blank_line > 0:
                        blank_line = 0

                    serial_count += 1
                    row[0].value = serial_count
                
            else:
                blank_line += 1
                if blank_line > 1:
                    rows_to_delete.append(row[0].row)

        # Delete marked rows
        for row in reversed(rows_to_delete):
            sheet.delete_rows(row)
            
        for row in sheet.iter_rows():
            for cell in row:
                cell.border = no_border
        
        # Save modified Excel
        # st.write(input_file)
        if not isinstance(input_file, str):
            file_name, file_extension = os.path.splitext(input_file.name)
        else:
            file_name, file_extension = os.path.splitext(input_file)
        output_file = f"{file_name}_modified{file_extension}"
        workbook.save(output_file)

        # Convert modified Excel to PDF
        excel_file = os.path.join(os.getcwd(), output_file)
        pdf_file = os.path.join(os.getcwd(), f"{file_name}_modified.pdf")
        convert_excel_to_pdf(excel_file, pdf_file)

        return pdf_file

def convert_excel_to_pdf(excel_file, pdf_file):
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    workbook = excel.Workbooks.Open(excel_file)

    sheet = workbook.Sheets(1)
      
    try:
        sheet.Select()
        workbook.ExportAsFixedFormat(0, pdf_file)
            
    except Exception as e:
        st.write(f"Error: {e}")
    finally:
        workbook.Close(False)
        excel.Quit()
        pythoncom.CoUninitialize()

def read_company_details(input_file):
    workbook = load_workbook(input_file)
    sheet = workbook.active  # Assuming data is in the first sheet

    company_details = {}
    for row in sheet.iter_rows(min_row=2, min_col=1, values_only=True):  # Assuming company details start from the second row, first column
        if row[0]:  # Assuming company names are in the first column
            company_name = row[0]
            email = row[1]
            poc = row[2]
            company_details[company_name] = {
                'Email': email,
                'POC': poc
            }

    return company_details

def fetch_products(file,date):
    # print("hi")
    
    workbook = load_workbook(file)
    
    sheet_names = workbook.sheetnames 
    # st.write(f"Sheet names: {sheet_names}")

    if len(sheet_names) > 1:
        for sheet_name in sheet_names[1:]:
            del workbook[sheet_name]

    sheet_names = workbook.sheetnames 
    # st.write(f"Remaining sheet names: {sheet_names}")

    sheet = workbook[sheet_names[0]]

    find_qty = False
    qty_column = None
    serial_count = 0
    blank_line = 0
    find_date=False
    date_column=None

    st.write(date)

    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None and isinstance(cell.value, str) and "date" in cell.value.strip().lower():
                find_date = True
                date_column = cell.column+1 
                old_date = sheet.cell(row=cell.row, column=date_column)
                # st.write(f"'qty' found at Row: {cell.row}, Column: {cell.column}")
                break
        if find_date:
            break

    # Find 'Qty.' column
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None and isinstance(cell.value, str) and cell.value.strip().lower() == "qty.":
                find_qty = True
                qty_column = cell.column 
                # st.write(f"'qty' found at Row: {cell.row}, Column: {cell.column}")
                break
        if find_qty:
            break

    if qty_column:
        rows_to_delete = []
        rows_product=[]

        # Modify rows based on conditions
        for row in sheet.iter_rows(min_row=cell.row + 1):
            qty_cell = row[qty_column - 1] 
            
            if row[1].value not in (None, '') or row[0].value not in (None, ''):
                if not isinstance(row[0].value, str): 
                    # rows_product.append([cell.value for cell in row[:5]])  # Assuming first 5 columns
                    # st.write(rows_product[-1])
                    # st.write(row[0].row)
                    row_data = {
                        "exel row no.":row[0].row,
                        "Serial No.": len(rows_product),
                        "Name": row[1].value,
                        "Qty.": row[qty_column - 1].value,
                        "Unit": row[3].value,
                        "Code": row[4].value
                    }
                    rows_product.append(row_data)
                    rows_product[-1]=row_data = {
                        "exel row no.":row[0].row,
                        "Serial No.": len(rows_product),
                        "Name": row[1].value,
                        "Qty.": row[qty_column - 1].value,
                        "Unit": row[3].value,
                        "Code": row[4].value
                    }
                    # st.write(rows_product)
                    if not isinstance(qty_cell.value, int):
                        rows_product[-1]=row_data = {
                        "exel row no.":row[0].row,
                        "Serial No.": len(rows_product),
                        "Name": row[1].value,
                        "Qty.": 0,
                        "Unit": row[3].value,
                        "Code": row[4].value
                    }
                        # if not isinstance(row[0].value, str):
                        rows_to_delete.append(row[0].row) 
                    else:
                        if blank_line > 0:
                            blank_line = 0

                        serial_count += 1
                        row[0].value = serial_count
                # st.write(rows_product[-1])
            else:
                blank_line += 1
                if blank_line > 1:
                    rows_to_delete.append(row[0].row)
        
        st.write("Product Table")
        if rows_product:
            df = pd.DataFrame(rows_product)
            column_config = {
            "Serial No.": st.column_config.Column("Serial No.",width=40),
            "Name": st.column_config.Column("Name",width=160),
            "Qty.": st.column_config.Column("Quantity",width=40),
            "Unit": st.column_config.Column("Unit",width=60),
            "Code": st.column_config.Column("Code")
            }
            
            edited_df = st.data_editor(df, column_config=column_config, hide_index=True,width=1500,column_order=("Serial No.","Name","Qty.","Unit","Code"))
            # edited_df = st.data_editor(df,disabled=())

            # Save button
            if st.button("Save Changes"):
                df.update(edited_df)

                # Optionally, you can save df back to Excel or perform other operations
                # st.write("Updated DataFrame:")
                # st.write(df)
                # st.write(rows_product)
                                    
                for index, valuee in df.iterrows():
                    sheet.cell(row=valuee["exel row no."], column=qty_column-1,value=valuee["Name"])
                    sheet.cell(row=valuee["exel row no."], column=qty_column,value=valuee["Qty."])
                    sheet.cell(row=valuee["exel row no."], column=qty_column+1, value=valuee["Unit"])
                    sheet.cell(row=valuee["exel row no."], column=qty_column+2, value=valuee["Code"])

                # # Save changes to the workbook
                
                # st.write(old_date.value)
                # st.write(date)
                # date_obj = datetime.strptime(date, "%Y-%m-%d")

                # # Format to the desired output
                # formatted_date = date_obj.strftime("%d/%m/%Y")
                # st.write(formatted_date)
                
                old_date.value=date
                st.success("SAVED SUCCESSFULLY")
                workbook.save(file)
                
                # st.write(file)
                pdf_path=modify_excel_and_convert_to_pdf(file)
                if pdf_path:
                    with open(pdf_path, "rb") as f:
                        base64_pdf = base64.b64encode(f.read()).decode('utf-8')
                        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="700" height="500" type="application/pdf"></iframe>'
                    st.markdown(pdf_display, unsafe_allow_html=True)
                
                    if st.button("Send PDF via Email"):
                        send_email("hi", pdf_path)
                        st.success("Email sent successfully!")
                        st.sidebar.empty()
                        st.sidebar.success("Email sent successfully!")
                
            

        else:
            st.info("No data to display.")
            

def send_email(recipient_email, pdf_file_path):
    st.write("PDF")
    

def main():
    st.title("Excel Company Selector")
    on_uploaded_file = st.sidebar.file_uploader("Upload an Excel file", type=["xlsx"])

    st.sidebar.title("Company Options")

    # Path to the pre-existing Excel file
    uploaded_file = 'com_list_test.xlsx'

    # st.sidebar.write("Using pre-existing file:")
    # st.sidebar.write(uploaded_file)
    company_details = read_company_details(uploaded_file)
    
    if not on_uploaded_file:
        if company_details:
            selected_company = st.sidebar.selectbox("Select a company", list(company_details.keys()))

            input_exel_file=f"{selected_company}.xlsx".strip().lower()
            # st.sidebar.write(f"You selected: {selected_company}")

            # Date selection
            selected_date = st.sidebar.date_input("Order date", date.today())

            # Display email addresses for the selected company
            if selected_company in company_details:
                emails = company_details[selected_company]['Email'].split(', ')  # Assuming emails are comma-separated
                selected_emails = st.sidebar.multiselect("Select email(s)", emails, default=emails)

                # st.sidebar.write(f"Selected email(s): {selected_emails}")

                st.sidebar.write(f"**POC:** {company_details[selected_company]['POC']}")
                
                pdf_file_path=fetch_products(input_exel_file, selected_date)
                # st.write(pdf_file_path)
                # if pdf_file_path:
                #     with open(pdf_file_path, "rb") as f:
                #         base64_pdf = base64.b64encode(f.read()).decode('utf-8')
                #         pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="700" height="500" type="application/pdf"></iframe>'
                #     st.markdown(pdf_display, unsafe_allow_html=True)
                    
                #     if st.button("Send PDF via Email"):
                #         send_email(selected_emails, pdf_file_path)
                #         st.success("Email sent successfully!")
                #         st.sidebar.empty()
                #         st.sidebar.success("Email sent successfully!")
                

        # Modify Excel and convert to PDF
    if on_uploaded_file:
        print(on_uploaded_file)
        pdf_file_path = modify_excel_and_convert_to_pdf(on_uploaded_file)
        print(pdf_file_path)      
        if pdf_file_path:
            with open(pdf_file_path, "rb") as f:
                base64_pdf = base64.b64encode(f.read()).decode('utf-8')
                pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="700" height="500" type="application/pdf"></iframe>'
            st.markdown(pdf_display, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
