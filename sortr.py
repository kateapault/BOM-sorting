# sortr.py takes a text file of the bullshit copy and paste track segments and puts it in excel so i don't have to copy pasta eight million damn things
import xlwt
from xlwt import Workbook

# input text file
textfile_name = input("Please enter name of text file (no extension): ")
info = open(textfile_name+".txt","r")
layout = info.readlines()

# input name for excel sheet
filename = input("Please enter what you want the excel file to be named (no extension): ")
excel_name = filename + ".xls"
# create excel workbook

wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')
# read text file

sheet1.write(0, 0, "Part Number")
sheet1.write(0, 1, "Part Description")
sheet1.write(0, 2, "Qty")
linenum = 1

# for each line, check if "ZFA" is in line
for line in layout:
# if so, separate line by spaces. third value will be qty, 4th will be PN, 5th will be part desc (first part)
    if "ZFA" in line:
        items = line.split()
# write PN, part desc, qty to excel
        sheet1.write(linenum, 0, items[3])
        sheet1.write(linenum, 1, items[4])
        sheet1.write(linenum, 2, items[2])
        linenum += 1
# else skip the line    
    else: pass

# save excel file
wb.save(excel_name)

#close text file
info.close()

# when done, popup: "All finished! Would you like to open the excel sheet?" yes/no button
# no - just close out
# yes - open excel sheet
 