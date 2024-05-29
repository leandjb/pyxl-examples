from openpyxl import load_workbook

wb = load_workbook('add_sheets.xlsx')

for sheet in wb:
    print(sheet.title)

    wb.cre

source = wb['transactions']
new_sheet = wb.copy_worksheet(source)
new_sheet.title = 'copied new title'

for sheet in wb:
    print(sheet.title)

wb.save('example3.xlsx')
