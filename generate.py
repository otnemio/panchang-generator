import calendar
from calendar import monthrange
import openpyxl

wb=openpyxl.load_workbook("hindi_cal_template.xlsx")
sheets=wb.sheetnames
sh2=wb['Sheet_Mapping']

mn=input("Month No.: ")
yr=input("Year: ")

sh2['E38']=calendar.month_name[int(mn)]

first_day=(monthrange(int(yr),int(mn))[0]+1)%7  #from 0[0-6]Monday to 0[0-6]Sunday
days=monthrange(int(yr),int(mn))[1]             #total no of days in the given month
for i in range(1,36):       #loop from 1 to 35
    if(first_day+i<=35):    #no need to worry
        sh2['E'+str(first_day+i)].value=i if i<=days else " "
        sh2['K'+str(first_day+i)].value=input("hindi day no. (1-15): ") if i<=days else "0"
        sh2['L'+str(first_day+i)].value=input("hindi day paksha (K or S): ") if i<=days else "E"
    else:                   #need to worry
        sh2['E'+str((first_day+i)%35)].value=i if i<=days else " "
        sh2['K'+str((first_day+i)%35)].value=input("hindi day no. (1-15): ") if i<=days else "0"
        sh2['L'+str((first_day+i)%35)].value=input("hindi day paksha (K or S): ") if i<=days else "E"


wb.save("hindi_cal_printable.xlsx")
