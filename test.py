import openpyxl
from openpyxl.styles import PatternFill
from main import minDistance

def Search(a,list):
    for i,x in enumerate(list):
        if minDistance(a,x)<3:
            return i
    return -1

workbook_read = openpyxl.load_workbook('test.xlsx')
worksheet_read = workbook_read['Full']
max_row=worksheet_read.max_row

my_achievements=[]
my_situation=[]
for i in range(1,max_row):
    my_achievements.append(worksheet_read['c'+str(i+1)].value)
    my_situation.append('达成' if not '/' in worksheet_read['e'+str(i+1)].value else worksheet_read['e'+str(i+1)].value)
workbook_read = openpyxl.load_workbook('2.4.0成就_CHS.xlsx')
worksheet_read = workbook_read['Full']
max_row=worksheet_read.max_row
fill = PatternFill(start_color ='FFFF00', end_color = 'FFFF00', fill_type = 'solid')

for i in range(1,max_row):
    achievement=worksheet_read['b'+str(i+1)].value
    index=Search(achievement, my_achievements)
    if index==-1 or (index!=-1 and my_situation[index]!='达成'):
        worksheet_read['b'+str(i+1)].fill=fill 

workbook_read.save('compare_ans.xlsx')

