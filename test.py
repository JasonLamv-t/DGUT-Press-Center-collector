import xlwt

workbook = xlwt.Workbook()
sheet = workbook.add_sheet('name', cell_overwrite_ok=True)

sheet.write()