import openpyxl
import json
from dateutil.rrule import *
from datetime import datetime
import calendar
import time

def scheduling():
    current_time=datetime.now()
    workbook = openpyxl.load_workbook('Schedule.xlsx')
    worksheet = workbook[calendar.day_name[current_time.weekday()]]

    start_hour=9
    start_minute=15

    break_start_hour=12
    break_start_minute=15

    break_end_hour=13
    break_end_minute=00

    end_hour=16
    end_minute=00


    loop_start_time=datetime(datetime.today().year, datetime.today().month, datetime.today().day, hour=start_hour, minute=start_minute, second=0) 
    break_start_time=datetime(datetime.today().year, datetime.today().month, datetime.today().day, hour=break_start_hour, minute=break_start_minute, second=0)  
    break_end_time=datetime(datetime.today().year, datetime.today().month, datetime.today().day, hour=break_end_hour, minute=break_end_minute, second=0) 
    end_time=datetime(datetime.today().year, datetime.today().month, datetime.today().day, hour=end_hour, minute=end_minute, second=0) 
   

    last=2
    for time_interval in rrule(freq=HOURLY,until=break_start_time,dtstart=loop_start_time):  
      if (current_time.hour>time_interval.hour  or(current_time.hour==time_interval.hour and current_time.minute>time_interval.minute)): 
            last=last+1
      else:
            break
    last=last+1  

    for time_interval in rrule(freq=HOURLY,until=end_time,dtstart=break_end_time):  
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
                    status.update({row[1].value:"+"})
                elif row[last-1].value=="No": 
                    status.update({row[1].value:"-"})  
                elif row[last-1].value==None:
                    break
                
    status_file=open("status.json",mode="w+")
    json.dump(status,status_file)
    
while(True):
    for a in range(0,4):
       scheduling()
       time.sleep(3600)
    scheduling()   
    time.sleep(2700)
    for a in range(0,4):
       scheduling()
       time.sleep(3600)



