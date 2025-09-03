# -*- coding: utf-8 -*-
# Importing libraries
import pandas as pd
from datetime import datetime, timedelta
import uuid
import re

'''
Lectures are recurring weekly with some exceptions stated in the schedule in their names('Аналитическая химия 2.09, 16.09, 30.09'). 
Columns are: День недели, № пары, Время, Ауд., Вид уч. занятий, Преподаватель, Нечетная неделя, Четная неделя, Преподаватель, Вид уч. занятий, Ауд.,  Время, № пары
This allows to determine the day of the week, time, type of class, location, instructor, and whether it occurs on odd or even weeks.
Which allows to create a calendar with recurring events with RRULES.
The semester starts on 01.09.2025 and ends on 21.12.2025 basically the end is UNTIL. 
All the times are in local time (UTC+3).

The schedule is divided into 6 parts, each part contains 17 rows of data
df=pd.read_excel(filename, header=None, sheet_name=student_group, usecols='B:H', skiprows=15, nrows=17)
print(df.iloc[1,2])
print(type(df.iloc[1,2]))

for i in range(6):
    sr=15+i*17
    dfodd = pd.read_excel(filename, header=None, sheet_name=student_group, usecols='I:N', skiprows=sr, nrows=17) # Schedule is always in range B16:N117

Doing in for cycle for 6 parts of the schedule.
First we grab  campus from [0,0], then day of the week from [1,0]. Then we loop through the rows and for each row 
the task is to check if the cell in [n,6] is not NaN, then it means that there is a class at this time on odd weeks. 
Next we get class name from [n,6], professor from [n,5], class type from [n,4], location from [n,3], time from [n,2], class number from [n,1] converted to integer.
Then we fill in the template for the VEVENT, replacing the placeholders with the actual values.
Also there are cases where [n,6] contains something like 'Аналитическая химия 2.09, 16.09, 30.09', in this case we need to extract the dates and put classes on these dates only.
Before iterating i we go through dfeven in similar fashion, class name is in [n,8], professor in [n,9], class type in [n,10], location in [n,11], time in [n,12], class number in [n,13].

The template for VEVENT is as follows:

BEGIN:VEVENT
UID:PLACEHOLDER-UID-1
SUMMARY:PLACEHOLDER-TITLE
LOCATION:PLACEHOLDER-LOCATION
DTSTART:YYYYMMDDTHHMMSSZ
DTEND:YYYYMMDDTHHMMSSZ
RRULE:to be filled
DESCRIPTION:ClassType: PLACEHOLDER-TYPE\nProfessor: PLACEHOLDER-NAME\n
END:VEVENT


Rules for filling in the placeholders:
UID: PLACEHOLDER-UID-1 is filled with a unique identifier generated using uuid.uuid4()
SUMMARY: PLACEHOLDER-TITLE is filled with class name
LOCATION: PLACEHOLDER-LOCATION is filled with location and campus
DTSTART and DTEND are filled with the start and end time of the class, in the format YYYYMMDDTHHMMSSZ
RRULE must state that event recurs once in two weeks, on a certain day of the week, until the end of the semester.
and UNTIL is filled with the end date of the semester in the format YYYYMMDDT235959 (20251221T235959)
DESCRIPTION is filled with class type and professor name
'''
# Set the variables for reading the Excel file, can be changed as needed
filename= 'ИХТиПЭ-очно-2курс.xlsx'                                                            # Filename for the downloaded Excel file
student_group= 'ХПУ-124'                                                                      # Student group which is also a sheet name in the Excel file
# Convert Russian to English day abbreviations
ru_days= {
    'ПН': 'MO',
    'ВТ': 'TU',
    'СР': 'WE',
    'ЧТ': 'TH',
    'ПТ': 'FR',
    'СБ': 'SA'}

with open(f'{student_group}_расписание.ics', 'w', encoding='utf-8') as f:
    f.write('BEGIN:VCALENDAR\n')
    f.write('VERSION:2.0\n')
    f.write('PRODID:-//k0rte5/Конвертер расписания в iCalendar//RU\n')
    for i in range(6):
        sr= 15+i*17
        dfodd= pd.read_excel(filename, header=None, dtype=str, sheet_name=student_group, usecols='B:H', skiprows=sr, nrows=17)
        dfeven= pd.read_excel(filename, header=None, dtype=str, sheet_name=student_group, usecols='I:N', skiprows=sr, nrows=17)
        campus= dfodd.iloc[0,0]
        weekday= ru_days[dfodd.iloc[1,0]]
        for j in range(1,17):
            if not pd.isna(dfodd.iloc[j,6]):
                if not re.search(r"\b\d{1,2}\.\d{1,2}\b", dfodd.iloc[j,6]):
                    class_name= dfodd.iloc[j,6]
                    professor= dfodd.iloc[j,5]
                    class_type= dfodd.iloc[j,4]
                    class_location= dfodd.iloc[j,3]
                    class_time= dfodd.iloc[j,2]
                    class_number= dfodd.iloc[j,1]

                    # Take start and end time from class_time
                    start_time= datetime.strptime(class_time.split('-')[0], '%H:%M')
                    end_time= datetime.strptime(class_time.split('-')[1], '%H:%M')
                    # Convert to UTC time (subtract 3 hours)
                    start_time_utc= (start_time - timedelta(hours=3)).strftime('%H%M%S')
                    end_time_utc= (end_time - timedelta(hours=3)).strftime('%H%M%S')
                    
                    f.write('BEGIN:VEVENT\n')
                    f.write(f'UID:{uuid.uuid4()}@rgukics\n')
                    f.write(f'SUMMARY:{class_name}\n')
                    f.write(f'LOCATION:{class_location}, {campus}\n')
                    f.write(f'DTSTART:2025090{i+1}T{start_time_utc}Z\n')
                    f.write(f'DTEND:2025090{i+1}T{end_time_utc}Z\n')
                    f.write(f'RRULE:FREQ=WEEKLY;INTERVAL=2;BYDAY={weekday};UNTIL=20251221T235959Z\n')
                    f.write(f'DESCRIPTION:{class_type}\\n{professor}\\n\n')
                    f.write('END:VEVENT\n')
                elif re.search(r"\b\d{1,2}\.\d{1,2}\b", dfodd.iloc[j,6]):
                    dates= re.findall(r"\b\d{1,2}\.\d{1,2}\b", dfodd.iloc[j,6])
                    class_name= re.sub(r"\s*\b\d{1,2}\.\d{1,2}\b(?:,\s*\b\d{1,2}\.\d{1,2}\b)*", '', dfodd.iloc[j,6]).strip()
                    professor= dfodd.iloc[j,5]
                    class_type= dfodd.iloc[j,4]
                    class_location= dfodd.iloc[j,3]
                    class_time= dfodd.iloc[j,2]
                    class_number= dfodd.iloc[j,1]

                    # Take start and end time from class_time
                    start_time= datetime.strptime(class_time.split('-')[0], '%H:%M')
                    end_time= datetime.strptime(class_time.split('-')[1], '%H:%M')
                    # Convert to UTC time (subtract 3 hours)
                    start_time_utc= (start_time - timedelta(hours=3)).strftime('%H%M%S')
                    end_time_utc= (end_time - timedelta(hours=3)).strftime('%H%M%S')

                    for date in dates:
                        day, month= map(int, date.split('.'))
                        event_date= datetime(2025, month, day).strftime('%Y%m%d')
                        f.write('BEGIN:VEVENT\n')
                        f.write(f'UID:{uuid.uuid4()}@rgukics\n')
                        f.write(f'SUMMARY:{class_name}\n')
                        f.write(f'LOCATION:{class_location}, {campus}\n')
                        f.write(f'DTSTART:{event_date}T{start_time_utc}Z\n')
                        f.write(f'DTEND:{event_date}T{end_time_utc}Z\n')
                        f.write(f'DESCRIPTION:{class_type}\\n{professor}\\n\n')
                        f.write('END:VEVENT\n')
            if not pd.isna(dfeven.iloc[j,0]):
                if not re.search(r"\b\d{1,2}\.\d{1,2}\b", dfeven.iloc[j,0]):
                    class_name= dfeven.iloc[j,0]
                    professor= dfeven.iloc[j,1]
                    class_type= dfeven.iloc[j,2]
                    class_location= dfeven.iloc[j,3]
                    class_time= dfeven.iloc[j,4]
                    class_number= dfeven.iloc[j,5]

                    # Take start and end time from class_time
                    start_time= datetime.strptime(class_time.split('-')[0], '%H:%M')
                    end_time= datetime.strptime(class_time.split('-')[1], '%H:%M')
                    # Convert to UTC time (subtract 3 hours)
                    start_time_utc= (start_time - timedelta(hours=3)).strftime('%H%M%S')
                    end_time_utc= (end_time - timedelta(hours=3)).strftime('%H%M%S')
                    
                    f.write('BEGIN:VEVENT\n')
                    f.write(f'UID:{uuid.uuid4()}@rgukics\n')
                    f.write(f'SUMMARY:{class_name}\n')
                    f.write(f'LOCATION:{class_location}, {campus}\n')
                    f.write(f'DTSTART:2025090{i+8}T{start_time_utc}Z\n')
                    f.write(f'DTEND:2025090{i+8}T{end_time_utc}Z\n')
                    f.write(f'RRULE:FREQ=WEEKLY;INTERVAL=2;BYDAY={weekday};UNTIL=20251221T235959Z\n')
                    f.write(f'DESCRIPTION:{class_type}\\n{professor}\\n\n')
                    f.write('END:VEVENT\n')
                elif re.search(r"\b\d{1,2}\.\d{1,2}\b", dfeven.iloc[j,0]):
                    dates= re.findall(r"\b\d{1,2}\.\d{1,2}\b", dfeven.iloc[j,0])
                    class_name= re.sub(r"\s*\b\d{1,2}\.\d{1,2}\b(?:,\s*\b\d{1,2}\.\d{1,2}\b)*", '', dfeven.iloc[j,0]).strip()
                    professor= dfeven.iloc[j,1]
                    class_type= dfeven.iloc[j,2]
                    class_location= dfeven.iloc[j,3]
                    class_time= dfeven.iloc[j,4]
                    class_number= dfeven.iloc[j,5]
                    # Take start and end time from class_time
                    start_time= datetime.strptime(class_time.split('-')[0], '%H:%M')
                    end_time= datetime.strptime(class_time.split('-')[1], '%H:%M')
                    # Convert to UTC time (subtract 3 hours)
                    start_time_utc= (start_time - timedelta(hours=3)).strftime('%H%M%S')
                    end_time_utc= (end_time - timedelta(hours=3)).strftime('%H%M%S')
                    for date in dates:
                        day, month= map(int, date.split('.'))
                        event_date= datetime(2025, month, day).strftime('%Y%m%d')
                        f.write('BEGIN:VEVENT\n')
                        f.write(f'UID:{uuid.uuid4()}@rgukics\n')
                        f.write(f'SUMMARY:{class_name}\n')
                        f.write(f'LOCATION:{class_location}, {campus}\n')
                        f.write(f'DTSTART:{event_date}T{start_time_utc}Z\n')
                        f.write(f'DTEND:{event_date}T{end_time_utc}Z\n')
                        f.write(f'DESCRIPTION:{class_type}\\n{professor}\\n\n')
                        f.write('END:VEVENT\n')
    f.write('END:VCALENDAR\n')