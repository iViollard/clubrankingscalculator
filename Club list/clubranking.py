#import the writer
import xlwt
#import the reader
import xlrd
#import the requests module
import requests
#go to the British Fencing Association website and download this file (specified)
url = "https://www.britishfencing.com/wp-content/uploads/2018/10/mf_oct_2018.xls"
downloaded_file = requests.get(url)
#write the contents to a new file called rankings.xls
with open("rankings.xls", 'wb') as file:
    file.write(downloaded_file.content)

#open the rankings spreadsheet
book = xlrd.open_workbook('rankings.xls')
#open the first sheet
first_sheet = book.sheet_by_index(0)
#print the values in the second column of the first sheet
print first_sheet.col_values(1)


#open a new spreadsheet
workbook = xlwt.Workbook()
#add a sheet named "Club BFA ranking"
worksheet1 = workbook.add_sheet("Club BFA ranking")
#in cell 0,0 (first cell of the first row) write "Ranking"
worksheet1.write(0, 0, "Ranking")
#in cell 0,1 (second cell of the first row) write "Name"
worksheet1.write(0, 1, "Name")  
#in cell 0,2 (third cell of the first row) write "NIF"
worksheet1.write(0, 2, "NIF")
#in cell 0,1 (fourth cell of the first row) write "Points"
worksheet1.write(0, 3, "Points") 
#in cell 0,1 (fifth cell of the first row) write "Home country"
worksheet1.write(0, 4, "Home country") 
#save and create the spreadsheet file
workbook.save("saxons.xls")

name = []
rank = []
NIF = []
points = []
hc = []
for i in range(first_sheet.nrows):
    #print(first_sheet.cell_value(i,3)) 
    if('Saxon' in first_sheet.cell_value(i,3)):  
        name.append(first_sheet.cell_value(i,1))
        rank.append(first_sheet.cell_value(i,8))
        NIF.append (first_sheet.cell_value(i,9))
        points.append (first_sheet.cell_value(i,10))
        hc.append (first_sheet.cell_value(i,6))
        print('a')
for j in range(len(name)):
    worksheet1.write(j+1,0,rank[j])
    worksheet1.write(j+1,1,name[j])
    worksheet1.write(j+1,2,NIF[j])
    worksheet1.write(j+1,3,points[j])
    worksheet1.write(j+1,4,hc[j])
workbook.save("saxons.xls")