#!/usr/bin/python3

import sys
import os.path
from icalendar import Calendar
import csv
import datetime
from datetime import timedelta
from dateutil import relativedelta
from pytz import timezone

filename = sys.argv[1]
# TODO: use regex to get file extension (chars after last period), in case it's not exactly 3 chars.
file_extension = str(sys.argv[1])[-3:]
headers = ('OnDuty', 'Start Time', 'End Time')
estzone = timezone('Europe/Sofia')

class CalendarEvent:
    """Calendar event class"""
    onduty = ''
    start = ''
    end = ''

    def __init__(self, name):
        self.name = name

events = []


def open_cal():
    if os.path.isfile(filename):
        if file_extension == 'ics':
            print("Extracting events from file:", filename, "\n")
            f = open(sys.argv[1], 'rb')
            gcal = Calendar.from_ical(f.read())

            for component in gcal.walk():
                event = CalendarEvent("event")
                event.onduty = component.get('ATTENDEE')
                if hasattr(component.get('dtstart'), 'dt'):
                    event.start = component.get('dtstart').dt.astimezone(estzone)
                if hasattr(component.get('dtend'), 'dt'):
                    event.end = component.get('dtend').dt.astimezone(estzone)
                
                events.append(event)
            f.close()
        else:
            print("You entered ", filename, ". ")
            print(file_extension.upper(), " is not a valid file format. Looking for an ICS file.")
            exit(0)
    else:
        print("I can't find the file ", filename, ".")
        print("Please enter an ics file located in the same folder as this script.")
        exit(0)


def csv_write(icsfile):
    csvfile = icsfile[:-3] + "csv"
    try:
        with open(csvfile, 'w') as myfile:
            wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
            wr.writerow(headers)
            for event in sortedevents:
                values = (event.onduty, event.start, event.end)
                wr.writerow(values)
            print("Wrote to ", csvfile, "\n")
    except IOError:
        print("Could not open file! Please close Excel!")
        exit(0)

def filterNonWorkingHours(duty):
    isWorkingHours = duty.start.date().weekday() not in [5,6] and duty.start.hour >= 8 and duty.end.hour <= 20
    print(isWorkingHours, "\t| ", duty.start.date().weekday(), " | ", duty.start.hour, " | ", duty.end.hour, "\n")        
    return not isWorkingHours
    #endEst = estzone.localize(duty.start)
    #return duties

def filterSofia():
    peopleToFilter = ["jveli@vmware.com", "mapostolova@vmware.com", "gbalabanov@vmware.com", "tjecheva@vmware.com", "tivanov@vmware.com"]
    sofiaDuties = list(filter(lambda obj: obj.onduty in peopleToFilter, events))
    filterFrom  = datetime.date(2020, 12, 1)
    filterTo = filterFrom + relativedelta.relativedelta(months=1)
    sofiaDutiesSelectedMonth = list(filter(lambda obj: obj.start.date() >= filterFrom and obj.start.date() < filterTo, sofiaDuties))
    return list(filter(lambda obj: filterNonWorkingHours(obj), sofiaDutiesSelectedMonth))

def debug_event(class_name):
    print("Contents of ", class_name.name, ":")
    print(class_name.onduty)
    print(class_name.start)
    print(class_name.end, "\n")

open_cal()
sortedevents=filterSofia()
#sortedevents=sorted(events, key=lambda obj: obj.start) # Needed to sort events. They are not fully chronological in a Google Calendard export ...
csv_write(filename)
#debug_event(event)
