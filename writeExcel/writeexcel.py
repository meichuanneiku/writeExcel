# import xlwings as xw
import os

def write_excel_xls_append(self, path, xlstitle, value, excelApp):
    if not os.path.exists(path):
        self.write_excel_xls(path, xlstitle, excelApp)

    wb = excelApp.books.open(path)
    sheet = wb.sheets[0]

    info = sheet.used_range

    crrentRows = info.last_cell.row
    row_str = str(crrentRows + 1)

    sheet.range('A' + row_str).value = value

    wb.save()
    wb.close()

def write_excel_xls(self, path, value, excelApp):
    wb = excelApp.books.add()
    sheet = wb.sheets[0]
    sheet.range('A1').value = value
    wb.save(path)
    wb.close()
