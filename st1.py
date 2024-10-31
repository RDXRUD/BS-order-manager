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
from openpyxl.styles import Font, PatternFill,Alignment
from openpyxl.utils import get_column_letter

def modify_excel_and_convert_to_pdf(input_file,date,company):
    save_directory=r"D:\PROJECTS\BS file manager\BS-order-manager\data files\Modified Files"

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
            if cell.value is not None and isinstance(cell.value, str) and (cell.value.strip().lower() == "qty." or cell.value.strip().lower() == "qty"):
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
            file_name, file_extension = os.path.splitext(company)
        
        output_directory = os.path.join(save_directory, str(date))
        os.makedirs(output_directory, exist_ok=True)
        st.write(output_directory)
        output_file = os.path.join(output_directory, f'order_{file_name}_{date}{file_extension}')
        
        # output_file = f"{save_directory}\{date}\order_{file_name}_{date}{file_extension}"
        # os.makedirs(output_file, exist_ok=True)
        st.write(output_file)
        workbook.save(output_file)

        # Convert modified Excel to PDF
        st.write("ee:")
        st.write(company)
        excel_file = os.path.join(os.getcwd(), output_file)
        pdf_file = os.path.join(os.getcwd(), f"{save_directory}\{date}\order_{file_name}_{date}.pdf")
        convert_excel_to_pdf(excel_file, pdf_file)

        return pdf_file

def convert_excel_to_pdf(excel_file, pdf_file):
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    workbook = excel.Workbooks.Open(excel_file)
    sheet = workbook.Sheets(1)
    
    try:
        # Adjust page setup to fit content to PDF
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesWide = 1  # Fit all columns to one page width
        sheet.PageSetup.FitToPagesTall = False  # Allow multiple pages for rows if needed
        
        sheet.Select()
        workbook.ExportAsFixedFormat(0, pdf_file)
            
    except Exception as e:
        print(f"Error: {e}")
    finally:
        workbook.Close(False)
        excel.Quit()
        pythoncom.CoUninitialize()

def read_company_details(input_file, file):
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

def fetch_products(file,date,company,tc,tl,orderNO,emails):
    # print("hi")
    name,extension=company.split(".")
    workbook = load_workbook(file)
    
    sheet_names = workbook.sheetnames 
    # st.write(sheet_names)
    # st.write(f"Sheet names: {sheet_names}")

    if len(sheet_names) > 1:
        for sheet_name in sheet_names[1:]:
            del workbook[sheet_name]

    sheet_names = workbook.sheetnames 
    # st.write(f"Remaining sheet names: {sheet_names}")

    sheet = workbook[sheet_names[0]]

    findHead=False
    find_qty = False
    qty_index = None
    serial_count = 0
    blank_line = 0
    find_date=False
    date_column=None
    headingRow=None
    name_index=None
    countStr=None
    findStar=False
    
    head=[]

    # st.write(date)

    # for row in sheet.iter_rows():
    #     for cell in row:
    #         if cell.value is not None and isinstance(cell.value, str) and "date" in cell.value.strip().lower():
    #             find_date = True
    #             date_column = cell.column+1 
    #             old_date = sheet.cell(row=cell.row, column=date_column)
    #             # st.write(f"'qty' found at Row: {cell.row}, Column: {cell.column}")
    #             break
    #     if find_date:
    #         break

    # Find 'Qty.' column
    for row in sheet.iter_rows():
        for cell in row:
            # st.write(cell.value)
            if cell.value is not None and isinstance(cell.value, str) and ("qty" in cell.value.strip().lower() or "name" in cell.value.strip().lower() or "size" in cell.value.strip().lower()):
                # st.write(cell.value)
                findHead=True
                
                # find_qty = True
                # qty_column = cell.column
                headingRow=cell.row 
                st.write(headingRow)
                # st.write(f"'qty' found at Row: {cell.row}, Column: {cell.column}")
                break
        if findHead:
            break
    
    if findHead:
        count=0
        for cell in sheet[headingRow]:
            st.write(cell.value)
            if (cell.value):

                if "qty" in cell.value.lower().strip() or "quantity" in cell.value.lower().strip():
                    find_qty=True
                    qty_index=count
                
                if "name" in cell.value.lower().strip() or "product" in cell.value.lower().strip():
                    name_index=count
                    
                st.write(cell.value)
                head.append(cell.value.strip())
                count+=1
                
        
    
        st.write(head)
        leng=len(head)
        # st.write(leng)
        blank=True
        # st.write(leng)
        product_rows=[]
        for row in sheet.iter_rows(min_row=headingRow + 1):
            row_data=[]
            for i in range (0,leng):

                row_data.append(row[i].value)
            product_rows.append(row_data) 
        
        # st.write(product_rows) 
        
        df = pd.DataFrame(product_rows,columns=head)
        df = df.dropna(how='all')
        for index, row in df.iterrows():
            if isinstance(row[head[0]], str) and all(pd.isna(row[col]) for col in head[1:]):
                df = df.drop(index)
        df[df.columns[0]] = range(1, len(df) + 1)
        st.write("ad:",df[df.columns[name_index]])
        countStr=1
        if "*" in str(df[df.columns[name_index]]):
            findStar=True
            for i in range(len(df)):
                if "*" in str(df.iloc[i, name_index]):
                    st.write("star found")
                    df.iloc[i, 0] = None
                    continue
                df.iloc[i, 0] = countStr
                countStr+=1
                st.write(df.iloc[i, name_index])


            
        df = df.reset_index(drop=True)
        df = df.replace('-', None)
        # Display the DataFrame
        # st.write(df)
        edited_df = st.data_editor(df, hide_index=True,width=1500)
        st.write(edited_df)
        if st.button("Save Changes"):
                edited_df = edited_df.dropna(axis=1, how='all').loc[:, (edited_df != 0).any(axis=0)]

                delete_indices = []
                stack = []  # temporary stack to store row indices of the current category
                quantity_found = False

                # Iterate through DataFrame
                for index, row in edited_df.iterrows():
                    # Detect start of a new category by checking 'Serial No.'
                    if pd.notna(row[edited_df.columns[0]]) :
                        # Check the previous category: if no quantity was found, mark all rows in the stack for deletion
                        if stack and not quantity_found:
                            delete_indices.extend(stack)
                        
                        # Reset for the new category
                        stack = [index]
                        quantity_found = False  # Reset quantity_found for new category
                    else:
                        # If it's a continuation of the current category, add index to stack
                        stack.append(index)

                    # If a quantity is found in this row, set quantity_found to True
                    if pd.notna(row[edited_df.columns[qty_index]]) and   row[edited_df.columns[qty_index]] != 0:
                        quantity_found = True

                # Final check for the last category
                if stack and not quantity_found:
                    delete_indices.extend(stack)

                # Drop rows where entire categories had no quantity
                edited_df.drop(delete_indices, inplace=True)

                # Reset index and print the updated DataFrame
                edited_df.reset_index(drop=True, inplace=True)
                st.write("Filtered DataFrame:")
                st.write(edited_df)
                
                if findStar:
                    countStr=1  
                    for i in range(len(edited_df)):
                        if "*" in str(edited_df.iloc[i, name_index]):
                            st.write("star found")
                            edited_df.iloc[i, 0] = None
                            continue
                        edited_df.iloc[i, 0] = countStr
                        countStr+=1
                        st.write(edited_df.iloc[i, name_index])
                
                excel_file = 'header_template.xlsx'  # Replace with your file path
                
                workbook = load_workbook(excel_file)

                # Assuming data is in the first sheet (index 0)
                sheet = workbook.worksheets[0]


                # Iterate through rows starting from row 15 (index 14 in Python)
                data_to_write = edited_df.values.tolist()
                font_style = Font(name='Arial', size=12, bold=True)
                fill_style = PatternFill(start_color='C4BD97', end_color='C4BD97', fill_type='solid')
                # align_style = Alignment(horizontal='center', vertical='center')
                align_left = Alignment(horizontal='left', vertical='center')
                # Write data starting from row 15
                for col_idx, header_value in enumerate(head, start=1):
                    cell = sheet.cell(row=15, column=col_idx)
                    cell.value = header_value
                    cell.font = font_style
                    cell.fill = fill_style
                    cell.alignment=align_left


                # Write data with formatting starting from row 15
                for row_idx, row_data in enumerate(data_to_write, start=17):
                    for col_idx, cell_value in enumerate(row_data, start=1):
                        cell = sheet.cell(row=row_idx, column=col_idx)
                        cell.value = cell_value
                        cell.font = font_style
                        cell.alignment=align_left
                
                cell = sheet.cell(row=row_idx+2, column=1)
                cell.value = "Thanking You.git"
                cell.font = font_style
                cell.alignment=align_left
                # for col in sheet.iter_cols(min_row=15, max_row=15):
                #     for cell in col:
                #         cell.alignment = align_style
                
                for col in sheet.columns:
                    max_length = 0
                    column = col[0].column_letter  # Get the column name
                    for cell in col[14:]:  # Start from row 15 (index 14)
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2) * 1.2  # Adjusted width formula
                    sheet.column_dimensions[column].width = adjusted_width
                
                sheet['G7']=date
                column_G_width = 15  # Adjust as needed based on your content and date display requirements
                sheet.column_dimensions['G'].width = column_G_width
                st.write(date)
                
                sheet['A8']=tc
                # column_G_width = 15  # Adjust as needed based on your content and date display requirements
                # sheet.column_dimensions['G'].width = column_G_width
                # st.write(date)
                
                sheet['A9']=tl
                sheet['G8']=orderNO
                sheet['A10']=emails
                # column_G_width = 15  # Adjust as needed based on your content and date display requirements
                # sheet.column_dimensions['G'].width = column_G_width
                # st.write(date)
                
                save_directory=r"D:\PROJECTS\BS file manager\BS-order-manager\data files\Modified Files"
                output_directory = os.path.join(save_directory, str(date))
                os.makedirs(output_directory, exist_ok=True)
                st.write(output_directory)
                output_file = os.path.join(output_directory, f'order_{name}_{date}.{extension}')
                pdf_file = os.path.join(os.getcwd(), f"{save_directory}\{date}\order_{name}_{date}.pdf")
                # Save the workbook
                workbook.save(output_file)
                convert_excel_to_pdf(output_file, pdf_file)
                if pdf_file:
                    with open(pdf_file, "rb") as f:
                        base64_pdf = base64.b64encode(f.read()).decode('utf-8')
                        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="700" height="500" type="application/pdf"></iframe>'
                    st.markdown(pdf_display, unsafe_allow_html=True)
                
                
                
#                 df.update(edited_df)

#                 # Optionally, you can save df back to Excel or perform other operations
#                 # st.write("Updated DataFrame:")
#                 # st.write(df)
#                 # st.write(rows_product)
            #     else:
            #         blank=True
            # if not blank:
                
         
    
            # if cell.value is not None and ("qty" in cell.value.strip().lower() or "quantity" in cell.value.strip().lower()):
            #     qty_column=cell.column
            #     find_qty=True
            # st.write(cell.value)
           
            
#     if qty_column:
#         rows_to_delete = []
#         rows_product=[]

#         # Modify rows based on conditions
#         for row in sheet.iter_rows(min_row=headingRow + 1):
#             qty_cell = row[qty_column - 1] 
            
#             if row[1].value not in (None, '') or row[0].value not in (None, ''):
#                 if not isinstance(row[0].value, str): 
#                     # rows_product.append([cell.value for cell in row[:5]])  # Assuming first 5 columns
#                     # st.write(rows_product[-1])
#                     # st.write(row[0].row)
#                     row_data = {
#                         "exel row no.":row[0].row,
#                         "Serial No.": len(rows_product),
#                         "Name": row[1].value,
#                         "Qty.": row[qty_column - 1].value,
#                         "Unit": row[3].value,
#                         "Code": row[4].value
#                     }
#                     rows_product.append(row_data)
#                     rows_product[-1]=row_data = {
#                         "exel row no.":row[0].row,
#                         "Serial No.": len(rows_product),
#                         "Name": row[1].value,
#                         "Qty.": row[qty_column - 1].value,
#                         "Unit": row[3].value,
#                         "Code": row[4].value
#                     }
#                     # st.write(rows_product)
#                     if not isinstance(qty_cell.value, int):
#                         rows_product[-1]=row_data = {
#                         "exel row no.":row[0].row,
#                         "Serial No.": len(rows_product),
#                         "Name": row[1].value,
#                         "Qty.": 0,
#                         "Unit": row[3].value,
#                         "Code": row[4].value
#                     }
#                         # if not isinstance(row[0].value, str):
#                         rows_to_delete.append(row[0].row) 
#                     else:
#                         if blank_line > 0:
#                             blank_line = 0

#                         serial_count += 1
#                         row[0].value = serial_count
#                 # st.write(rows_product[-1])
#             else:
#                 blank_line += 1
#                 if blank_line > 1:
#                     rows_to_delete.append(row[0].row)
        
#         st.write("Product Table")
#         if rows_product:
#             df = pd.DataFrame(rows_product)
#             column_config = {
#             "Serial No.": st.column_config.Column("Serial No.",width=40),
#             "Name": st.column_config.Column("Name",width=160),
#             "Qty.": st.column_config.Column("Quantity",width=40),
#             "Unit": st.column_config.Column("Unit",width=60),
#             "Code": st.column_config.Column("Code")
#             }
            
#             edited_df = st.data_editor(df, column_config=column_config, hide_index=True,width=1500,column_order=("Serial No.","Name","Qty.","Unit","Code"))
#             # edited_df = st.data_editor(df,disabled=())

#             # Save button
#             if st.button("Save Changes"):
#                 edited_df['Qty.'] = edited_df['Qty.'].fillna(0)
#                 edited_df['Code'] = edited_df['Code'].fillna("")
#                 # st.write(edited_df)
#                 df.update(edited_df)

#                 # Optionally, you can save df back to Excel or perform other operations
#                 # st.write("Updated DataFrame:")
#                 # st.write(df)
#                 # st.write(rows_product)
                                    
#                 for index, valuee in df.iterrows():
#                     sheet.cell(row=valuee["exel row no."], column=qty_column-1,value=valuee["Name"])
#                     sheet.cell(row=valuee["exel row no."], column=qty_column,value=valuee["Qty."])
#                     sheet.cell(row=valuee["exel row no."], column=qty_column+1, value=valuee["Unit"])
#                     sheet.cell(row=valuee["exel row no."], column=qty_column+2, value=valuee["Code"])

#                 # # Save changes to the workbook
                
#                 # st.write(old_date.value)
#                 # st.write(date)
#                 # date_obj = datetime.strptime(date, "%Y-%m-%d")

#                 # # Format to the desired output
#                 # formatted_date = date_obj.strftime("%d/%m/%Y")
#                 # st.write(formatted_date)
                
#                 old_date.value=date
#                 st.success("SAVED SUCCESSFULLY")
#                 workbook.save(file)
                
#                 # st.write(file)
#                 pdf_path=modify_excel_and_convert_to_pdf(file,date,company)
#                 if pdf_path:
#                     with open(pdf_path, "rb") as f:
#                         base64_pdf = base64.b64encode(f.read()).decode('utf-8')
#                         pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="700" height="500" type="application/pdf"></iframe>'
#                     st.markdown(pdf_display, unsafe_allow_html=True)
                
#                     if st.button("Send PDF via Email"):
#                         send_email("hi", pdf_path)
#                         st.success("Email sent successfully!")
#                         st.sidebar.empty()
#                         st.sidebar.success("Email sent successfully!")
                
            

#         else:
#             st.info("No data to display.")
            

# def send_email(recipient_email, pdf_file_path):
#     st.write("PDF")
    

def main():
    st.title("ORDER MANAGEMENT")
    
    st.sidebar.title("Company Options")
    directory = r'D:\PROJECTS\BS file manager\BS-order-manager\data files\Purchase Order'
    file_list = os.listdir(directory)
    # st.write(file_list)
    
    # st.sidebar.write("Using pre-existing file:")
    # st.sidebar.write(uploaded_file)
    # company_details = read_company_details(uploaded_file, file_list)
    
    # if not on_uploaded_file:
    #     if company_details:
    #         selected_company = st.sidebar.selectbox("Select a company", list(company_details.keys()))

    #         input_exel_file=f"{selected_company}.xlsx".strip().lower()
    #         # st.sidebar.write(f"You selected: {selected_company}")

    #         # Date selection
    #         selected_date = st.sidebar.date_input("Order date", date.today())

    #         # Display email addresses for the selected company
    #         if selected_company in company_details:
    #             emails = company_details[selected_company]['Email'].split(', ')  # Assuming emails are comma-separated
    #             selected_emails = st.sidebar.multiselect("Select email(s)", emails, default=emails)

    #             # st.sidebar.write(f"Selected email(s): {selected_emails}")

    #             st.sidebar.write(f"**POC:** {company_details[selected_company]['POC']}")
                
    #             pdf_file_path=fetch_products(input_exel_file, selected_date)
    #             # st.write(pdf_file_path)
    #             # if pdf_file_path:
    #             #     with open(pdf_file_path, "rb") as f:
    #             #         base64_pdf = base64.b64encode(f.read()).decode('utf-8')
    #             #         pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="700" height="500" type="application/pdf"></iframe>'
    #             #     st.markdown(pdf_display, unsafe_allow_html=True)
                    
    #             #     if st.button("Send PDF via Email"):
    #             #         send_email(selected_emails, pdf_file_path)
    #             #         st.success("Email sent successfully!")
    #             #         st.sidebar.empty()
    #             #         st.sidebar.success("Email sent successfully!")
                

# FROM DIRECTORY
    
    if file_list:
        orderNO=st.sidebar.text_input("Order Number", 1)
        selectedCompany = st.sidebar.selectbox("Select a company", file_list)
        # st.write(selected_company)
        # input_exel_file=f"{selected_company}.xlsx"
        # st.sidebar.write(f"You selected: {selected_company}")

        # Date selection
        selectedDate = st.sidebar.date_input("Order date", date.today())
        textCompany=st.sidebar.text_input("Company Name", selectedCompany.split(".")[0])
        textLocation=st.sidebar.text_input("Company Location", " ")
        emails=st.sidebar.text_input("Email Ids", " ")

        # Display email addresses for the selected company
        if selectedCompany in file_list:
            # emails = company_details[selected_company]['Email'].split(', ')  # Assuming emails are comma-separated
            # selected_emails = st.sidebar.multiselect("Select email(s)", emails, default=emails)

            # # st.sidebar.write(f"Selected email(s): {selected_emails}")

            # st.sidebar.write(f"**POC:** {company_details[selected_company]['POC']}")
            st.write(selectedCompany)
            filePath = os.path.join(directory, selectedCompany)
            st.write(filePath)

            pdf_file_path=fetch_products(filePath, selectedDate,selectedCompany,textCompany,textLocation,orderNO,emails)
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
    # if on_uploaded_file:
    #     print(on_uploaded_file)
    #     pdf_file_path = modify_excel_and_convert_to_pdf(on_uploaded_file)
    #     print(pdf_file_path)      
    #     if pdf_file_path:
    #         with open(pdf_file_path, "rb") as f:
    #             base64_pdf = base64.b64encode(f.read()).decode('utf-8')
    #             pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="700" height="500" type="application/pdf"></iframe>'
    #         st.markdown(pdf_display, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
