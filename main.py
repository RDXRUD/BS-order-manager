import os
from openpyxl import load_workbook
import asposecells as cells

input_file="BD.xlsx"
workbook = load_workbook(input_file)
sheet = workbook.active
find_qty = False
find_serial = False
qty_column = None
serial_column = None
serial_count=0
blank_line=0

file_name, file_extension = os.path.splitext(input_file)
output_file = f"{file_name}_modified{file_extension}"

for row in sheet.iter_rows():
    for cell in row:
        if cell.value is not None:
            if isinstance(cell.value, str) and cell.value.strip().lower() == "qty.":
                find_qty = True
                qty_column = cell.column 
                print(f"'qty' found at Row: {cell.row}, Column: {cell.column}")
                break
            
    if find_qty:
        break

if qty_column:
    rows_to_delete = []

    for row in sheet.iter_rows(min_row=cell.row + 1):
        qty_cell = row[qty_column - 1] 
        
        if row[1].value not in (None, '') or row[0].value not in (None, '') : 
            if not isinstance(qty_cell.value, int):
                if not isinstance(row[0].value,str):
                    print("row val:",row[0].value)
                    rows_to_delete.append(row[0].row) 
                
            else:
                if blank_line>0:
                    blank_line=0

                serial_count+=1
                row[0].value=serial_count
                
        else:
            blank_line+=1
            if blank_line>1:
                rows_to_delete.append(row[0].row)

    for row in reversed(rows_to_delete):
        sheet.delete_rows(row)
        
workbook.save(output_file)



def excel_to_pdf(excel_file_path, pdf_file_path):
    # Load the Excel file
    workbook = cells.Workbook(excel_file_path)

    # Save the workbook as a PDF
    workbook.save(pdf_file_path, cells.SaveFormat.PDF)

# Example usage
excel_file = 'output_file'
pdf_file = f"{file_name}.pdf"
excel_to_pdf(excel_file, pdf_file)
