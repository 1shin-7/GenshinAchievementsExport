#import openpyxl 

#workbook_read = openpyxl.load_workbook('2.4.0成就_CHS.xlsx')
#worksheet_read = workbook_read['Full']
#max_row=worksheet_read.max_row
    
#achievements=[]
#for i in range(1,max_row):
#    achievements.append(worksheet_read['B'+str(i+1)].value)

from main import Find 

print(Find('白之光'))