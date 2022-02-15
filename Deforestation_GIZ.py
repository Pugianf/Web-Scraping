import xlrd, xlwt
from xlutils.copy import copy

read_book = xlrd.open_workbook('C:/Users/felipe.jardim/Desktop/Felipe/Dados/Web-Scraping/Pasta1.xls') #Make Readable Copy
write_book = copy(read_book) #Make Writeable Copy

write_sheet1 = write_book.get_sheet(1) #Get sheet 1 in writeable copy
write_sheet1.write(1, 11, 'test') #Write 'test' to cell (B, 11)

write_sheet2 = write_book.get_sheet(2) #Get sheet 2 in writeable copy
write_sheet2.write(3, 14, '135') #Write '135' to cell (D, 14)

write_book.save('C:/Users/felipe.jardim/Desktop/Felipe/Dados/Web-Scraping/Pasta1.xls')
