#import the writer
import xlwt
#import the reader
import xlrd
#import the requests module
import requests

#go and get the latest ranking files from the BFA website

#mens foil
url = "https://www.britishfencing.com/wp-content/uploads/2018/12/mf_dec_2018.xls"
downloaded_file = requests.get(url)
#write the contents to a new file called mf_rankings.xls
with open("mf_rankings.xls", 'wb') as file:
    file.write(downloaded_file.content)
    
#womens foil
url = "https://www.britishfencing.com/wp-content/uploads/2018/12/wf_dec_2018.xls"
downloaded_file = requests.get(url)

#write the contents to a new file called wf_rankings.xls
with open("wf_rankings.xls", 'wb') as file:
    file.write(downloaded_file.content)

#open the mf rankings spreadsheet
book = xlrd.open_workbook('mf_rankings.xls')
#create two new sheets in the new file
first_sheet = book.sheet_by_index(0)
#print the values in the second column of the first sheet
print first_sheet.col_values(1)

#open the wf rankings spreadsheet
book2 = xlrd.open_workbook('wf_rankings.xls')
#create two new sheets in the new file
second_sheet = book2.sheet_by_index(0)
#print the values in the second column of the first sheet
print second_sheet.col_values(1)


#open a new file
workbook = xlwt.Workbook()

#add a sheet named "Club MF BFA ranking"
worksheet1 = workbook.add_sheet("Club MF BFA ranking")
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

#add a second sheet named "Club WF BFA ranking"
worksheet2 = workbook.add_sheet("Club WF BFA ranking")
#in cell 0,0 (first cell of the first row) write "Ranking"
worksheet2.write(0, 0, "Ranking")
#in cell 0,1 (second cell of the first row) write "Name"
worksheet2.write(0, 1, "Name")  
#in cell 0,2 (third cell of the first row) write "NIF"
worksheet2.write(0, 2, "NIF")
#in cell 0,1 (fourth cell of the first row) write "Points"
worksheet2.write(0, 3, "Points") 
#in cell 0,1 (fifth cell of the first row) write "Home country"
worksheet2.write(0, 4, "Home country") 

#save and create the spreadsheet file
workbook.save("saxons.xls")

name = []
rank = []
NIF = []
points = []
hc = []
club = "Saxon"

for i in range(first_sheet.nrows):
    if(club in first_sheet.cell_value(i,3)):  
        name.append(first_sheet.cell_value(i,1))
        rank.append(first_sheet.cell_value(i,0))
        NIF.append (first_sheet.cell_value(i,9))
        points.append (first_sheet.cell_value(i,10))
        hc.append (first_sheet.cell_value(i,6))
        print('a')
        
for i in range(second_sheet.nrows):
    if(club in second_sheet.cell_value(i,3)):  
        name.append(second_sheet.cell_value(i,1))
        rank.append(second_sheet.cell_value(i,0))
        NIF.append (second_sheet.cell_value(i,9))
        points.append (second_sheet.cell_value(i,10))
        hc.append (second_sheet.cell_value(i,6))
        print('b')
        
for j in range(len(name)):
    worksheet1.write(j+1,0,rank[j])
    worksheet1.write(j+1,1,name[j])
    worksheet1.write(j+1,2,NIF[j])
    worksheet1.write(j+1,3,points[j])
    worksheet1.write(j+1,4,hc[j])
    
for k in range(len(name)):
    
    worksheet2.write(k+1,0,rank[k])
    worksheet2.write(k+1,1,name[k])
    worksheet2.write(k+1,2,NIF[k])
    worksheet2.write(k+1,3,points[k])
    worksheet2.write(k+1,4,hc[k])
    
workbook.save("saxons.xls")