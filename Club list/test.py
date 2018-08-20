#import the writer
import xlwt
#open the spreadsheet
workbook = xlwt.Workbook()
#add a sheet named "Club BFA ranking"
worksheet1 = workbook.add_sheet("Club BFA ranking")
#in cell 0,0 (first cell of the first row) write "Ranking"
worksheet1.write(0, 0, "Ranking")
#in cell 0,1 (second cell of the first row) write "Name"
worksheet1.write(0, 1, "Name")
#save and create the spreadsheet file
workbook.save("saxons.xls")

#import the reader
import xlrd
#open the spreadsheet
book = xlrd.open_workbook('mf_mar_2018.xls')
#open the first sheet
first_sheet = book.sheet_by_index(0)
# print the values in the second column of the first sheet
print first_sheet.col_values(1)