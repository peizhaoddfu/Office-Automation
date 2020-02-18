import xlrd
xlsx = xlrd.open_workbook('/Users/tnjmytuu/Documents/tpz.xlsx')
table = xlsx.sheet_by_index(0)
print(table.cell_value(1,4))
print(table.cell(1,4).value)
print(table.row(1)[4].value)

import xlwt
new_workbook = xlwt.Workbook()
worksheet = new_workbook.add_sheet('sheet_test')
worksheet.write(0,0,'test')
new_workbook.save('/Users/tnjmytuu/Documents/test.xls')
