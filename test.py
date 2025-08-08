import xlwings as xw

wb = xw.Book("poročanje proizvodnje2025.xlsm")
wb.app.visible = False
wb.app.calculation = 'automatic'
wb.save()
wb.close()