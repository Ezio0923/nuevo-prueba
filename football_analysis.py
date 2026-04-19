import openpyxl
import json

# Load the workbook
workbook = openpyxl.load_workbook('football_analysis.xlsx')

# Prepare the data to be converted to JSON
sheets_data = {}

# Extract information from each sheet
for sheet in workbook.worksheets:
    sheet_name = sheet.title
    dimensions = sheet.dimensions
    headers = [cell.value for cell in sheet[1]]
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        data.append(row)
    sheets_data[sheet_name] = {
        'dimensions': dimensions,
        'headers': headers,
        'data': data
    }

# Convert to JSON format
json_data = json.dumps(sheets_data, indent=4)

# Write the JSON data to a file
output_file = 'football_analysis.json'
with open(output_file, 'w') as f:
    f.write(json_data)
