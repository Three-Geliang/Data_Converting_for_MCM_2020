import openpyxl
import re
import sys

data_file_name = sys.argv[1]
book_name = data_file_name.replace('.tsv.tmp', '.xlsx')
sheet_name = 'Sheet1'
line_start = 1


print('Reading ', data_file_name)

current_line_working = line_start
current_col_working = 1
last_item = ''

fo = open(data_file_name, 'r+')
wb = openpyxl.Workbook()

while True:
    line = fo.readline()
    if not line:
        break
    if line == ' ':
        pass
    elif line == '\n':
        pass
    else:
        data = line.strip('\n')
        for item in data.split('\t'):
            if item == '':
                pass
            else:
                
                if re.search(r'[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}', last_item) or last_item == 'review_date':
                    if item == 'US' or item == 'us':
                        current_line_working += 1
                        current_col_working = 1
                sheet = wb.active
                print(item)
                sheet.cell(row=current_line_working, column=current_col_working).value = item
                
                current_col_working += 1
                last_item = item

print('Writing to ', book_name)
wb.save(book_name)
