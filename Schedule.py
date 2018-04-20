import openpyxl
import json
from dateutil.rrule import *
from datetime import datetime
import calendar

#put sample time to test
# current_time=datetime(datetime.today().year, datetime.today().month, datetime.today().day,hour=10,minute=14)

#to get current time 
current_time=datetime.now()

workbook = openpyxl.load_workbook('Schedule.xlsx')
worksheet = workbook[calendar.day_name[current_time.weekday()]]

start_hour=10
start_minute=15

end_hour=22
end_minute=15


loop_start_time=datetime(datetime.today().year, datetime.today().month, datetime.today().day, hour=start_hour, minute=start_minute, second=0) 
loop_end_time=datetime(datetime.today().year, datetime.today().month, datetime.today().day, hour=end_hour, minute=end_minute, second=0) 
if (current_time.hour==start_hour and current_time.minute<start_minute) or (current_time.hour==end_hour and current_time.minute>end_minute):
            print("wrong time input") 
            exit()

last=2
for time_interval in rrule(freq=HOURLY,until=loop_end_time,dtstart=loop_start_time):  
  if (current_time.hour>time_interval.hour  or(current_time.hour==time_interval.hour and current_time.minute>time_interval.minute)): 
        last=last+1
  else:
        break

         
check_availability=json.load(open('availability.json'))
print(check_availability)
available=[index for index,status in check_availability.items() if status=="present"]
print(available)
status={}
print (last)
for row in worksheet.iter_rows(min_row=2, max_col=last, max_row=10,min_col=1):
        if(str(row[0].value) in  available):
            if row[last-1].value=="Yes":
                status[row[1].value]={"status":"Yes"}
            elif row[last-1].value=="No": 
                status[row[1].value]={"status":"No"}   
            elif row[last-1].value==None:
                break
            
status_file=open("status.json",mode="w+")
json.dump(status,status_file)       

