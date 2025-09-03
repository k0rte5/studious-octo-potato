# -*- coding: utf-8 -*-
# Importing libraries
import pandas as pd
from ics import Calendar, Event
from datetime import datetime, timedelta

'''
Lectures are recurring weekly with some exceptions stated in the schedule in their names('Аналитическая химия 2.09, 16.09, 30.09'). 
Columns are: День недели, № пары, Время, Ауд., Вид уч. занятий, Преподаватель, Нечетная неделя, Четная неделя, Преподаватель, Вид уч. занятий, Ауд.,  Время, № пары
This allows to determine the day of the week, time, type of class, location, instructor, and whether it occurs on odd or even weeks.
Which allows to create a calendar with recurring events with RRULES.
'''

