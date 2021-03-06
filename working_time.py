# -*- coding: utf-8 -*-
from dateutil import parser
import calendar
import configparser


config = configparser.ConfigParser()
config.read('settings.ini')

YEAR = int(config['DATES']['YEAR'])
HOLIDAYS = config['DATES']['HOLIDAYS']
HOLIDAYS = [i.strip(',') for i in HOLIDAYS.split()]
WORKING_DATES = config['DATES']['WORKING_DATES']
WORKING_DATES = [i.strip(',') for i in WORKING_DATES.split()]


def convert_to_seconds(hours, minutes):
    return hours * 3600 + minutes * 60


def compute_working_time(_start, _end, dayfirst=False, result_in_hours=True):
    end = parser.parse(_end, dayfirst=dayfirst)
    start = parser.parse(_start, dayfirst=dayfirst)
    result = 0
    for month in range(start.month, end.month + 1):
        days = end.day + 1
        first_day = start.day
        if month != end.month:
            days = calendar.monthrange(YEAR, month)[1] + 1
        if month != start.month:
            first_day = 1
        for day in range(first_day, days):
            next_date = str(day) + '.' + str(month) + '.' + str(YEAR)
            next_date = parser.parse(next_date, dayfirst=True).date()
            if ((next_date.weekday() < 5) or
               (next_date in [parser.parse(i, dayfirst=True).date() for i in WORKING_DATES])) and \
               (next_date not in [parser.parse(i, dayfirst=True).date() for i in HOLIDAYS]):
                start_hour = 0
                start_minute = 0
                end_hour = 24
                end_minute = 0
                if next_date == start.date():
                    if convert_to_seconds(start.hour, start.minute) >= \
                       convert_to_seconds(start_hour, start_minute):
                        start_hour = start.hour
                        start_minute = start.minute
                    if convert_to_seconds(start.hour, start.minute) >= \
                       convert_to_seconds(end_hour, end_minute):
                           start_hour = end_hour
                           start_minute = end_minute
                        
                if next_date == end.date():
                    if convert_to_seconds(end.hour, end.minute) <= \
                       convert_to_seconds(end_hour, end_minute):
                        end_hour = end.hour
                        end_minute = end.minute
                    if convert_to_seconds(end_hour, end_minute) <= \
                       convert_to_seconds(start_hour, start_minute):
                           start_hour = end_hour
                           start_minute = end_minute
                
                converted_start = convert_to_seconds(start_hour, start_minute)
                converted_end = convert_to_seconds(end_hour, end_minute)
                result += converted_end - converted_start
    if result_in_hours:
        result /= 3600
    return result
#
# time = compute_working_time('2019-08-19 09:10:00', '2019-08-26 09:00:00')
# print(time)
