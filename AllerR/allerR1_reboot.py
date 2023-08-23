import re
import xlsxwriter

# Read the input text file
input_file = "terminal_output.txt"

# Initialize Excel workbook and worksheet
output_file = "allerR1_data.xlsx"
workbook = xlsxwriter.Workbook(output_file)
worksheet = workbook.add_worksheet()

# Add column headers with bold and centered formatting
bold_centered_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
headers = ['Date', 'Trajet', 'Company', 'Price 1PAX']
for col, header in enumerate(headers):
    worksheet.write(0, col, header, bold_centered_format)

# Regular expression patterns to extract relevant data
date_pattern = re.compile(r'le: (\d{4}-\d{2}-\d{2}) \d{2}:\d{2}:\d{2}\.\d+')
trajet_pattern = re.compile(r'Trajet: (.+)')
company_pattern = re.compile(r'Company: (.+)')
price_pattern = re.compile(r'(\d+,\d+)')

# Add a dictionary to track colors for each unique date
date_colors = {}

# Initialize column widths
col_widths = [len(header) for header in headers]

# Read and process the input text file
row = 1
date = None
trajet = None
company = None
price = None

with open(input_file, 'r') as file:
    for line in file:
        # Search for relevant patterns in the line
        date_match = date_pattern.search(line)
        trajet_match = trajet_pattern.search(line)
        company_match = company_pattern.search(line)
        price_match = price_pattern.search(line)
        
        # Process line based on patterns
        if date_match:
            date = date_match.group(1)
            if date not in date_colors:
                color = (220, 230, 240)  # Define a light color (RGB values)
                date_colors[date] = workbook.add_format({'bg_color': f'#{color[0]:02X}{color[1]:02X}{color[2]:02X}'})
        
        # Apply color based on date
        row_format = date_colors.get(date, None)
        
        if trajet_match:
            trajet = trajet_match.group(1)
        elif company_match:
            if price_match:
                company = company_match.group(1)
                price = price_match.group(1)
                worksheet.write(row, 0, date, row_format)
                worksheet.write(row, 1, trajet, row_format)
                worksheet.write(row, 2, company, row_format)
                worksheet.write(row, 3, price, row_format)
                row += 1

                # Update column widths based on content length
                col_widths[0] = max(col_widths[0], len(date))
                col_widths[1] = max(col_widths[1], len(trajet))
                col_widths[2] = max(col_widths[2], len(company))
                col_widths[3] = max(col_widths[3], len(price))
            else:
                company = company_match.group(1)
                price = None
        elif price_match and company:
            price = price_match.group(1)
            worksheet.write(row, 0, date, row_format)
            worksheet.write(row, 1, trajet, row_format)
            worksheet.write(row, 2, company, row_format)
            worksheet.write(row, 3, price, row_format)
            row += 1

            # Update column widths based on content length
            col_widths[0] = max(col_widths[0], len(date))
            col_widths[1] = max(col_widths[1], len(trajet))
            col_widths[2] = max(col_widths[2], len(company))
            col_widths[3] = max(col_widths[3], len(price))
            company = None

# Set column widths
for col, width in enumerate(col_widths):
    worksheet.set_column(col, col, width + 2)  # Add a small buffer for padding

# Close the Excel workbook
workbook.close()

print(f"Excel file '{output_file}' created successfully.")
