import re
from openpyxl import Workbook


with open('Camilo-Carlota_Angela.md', 'r', encoding='utf-8') as file:
    text = file.read()

# Split the text by '#'
chapters = text.split('#')

# Create a new workbook and select the active worksheet
wb = Workbook()
ws = wb.active

# Add a header row
ws.append(['Chapter', 'Year'])

# Iterate through each chapter
for i, chapter in enumerate(chapters, start=1):
    # Find all years in the chapter
    years = re.findall(r'\b(1[0-9]{3}|2[0-9]{3})\b', chapter)
    
    # Add each year to the worksheet, convert year to integer
    for year in years:
        ws.append([i, int(year)])

# Save the workbook to a file
wb.save('Chapter_years_Camilo-Carlota_Angela.xlsx')