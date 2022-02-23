import openpyxl as op
import xlrd, xlwt
from xlutils.copy import copy

wb = op.load_workbook('Entwaldungszahlen_Jan_2022.xls')
ws = wb.active
print(ws["A2"])

read_book = xlrd.open_workbook('C:/Users/felipe.jardim/Desktop/Felipe/Dados/Web-Scraping/Entwaldungszahlen_Jan_2022.xls') #Make Readable Copy
write_book = copy(read_book) #Make Writeable Copy

write_sheet1 = write_book.get_sheet(1) #Get sheet 1 in writeable copy
write_sheet1.write(9, 7, 'test') #Write 'test' to cell (B, 11)

write_sheet2 = write_book.get_sheet(1) #Get sheet 2 in writeable copy
write_sheet2.write(9, 8, 135) #Write 135 to cell (D, 14)

write_book.save('C:/Users/felipe.jardim/Desktop/Felipe/Dados/Web-Scraping/Entwaldungszahlen_Jan_2022.xls')