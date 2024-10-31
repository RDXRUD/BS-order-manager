import openpyxl

# Load the Excel workbook
wb = openpyxl.load_workbook('bd.xlsx')

# Assuming you are working with the first sheet, you can specify the sheet name if needed
sheet = wb.active

# Initialize an empty string to store concatenated data
data_string = ""

# Iterate over rows and cells to concatenate data into a single string
for row in sheet.iter_rows():
    for cell in row:
        # Concatenate cell value to the data string, separated by a space for example
        if cell.value:
            data_string += str(cell.value) + " "

# Specify the path for the output text file
output_file = 'temp.txt'

# Write the concatenated data string to the text file
with open(output_file, 'w') as f:
    f.write(data_string)

# Close the workbook when done
wb.close()

print(f"Data saved to {output_file}")
