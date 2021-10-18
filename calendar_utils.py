import datetime
from datetime import timedelta

import numpy as np
import pandas as pd

from ics import Calendar, Event
import time


def read_from_excel(filename='test.xlsx'):
    df = pd.read_excel(filename)
    events = []
    for index, row in df.iterrows():
        if not pd.isnull(row['end']):
            row.fillna('', inplace=True)
            e = create_event(row['name'], row['begin'], end=row['end'], location=row['location'],
                             description=row['description'])
        else:
            row.fillna('', inplace=True)
            e = create_event(row['name'], row['begin'],
                             duration=timedelta(minutes=int(row['duration (min)'])), location=row['location'],
                             description=row['description'])
        events.append(e)
    return events


def create_event(name, begin, duration=timedelta(hours=0.5), end=None, location=None, description=None):
    event = Event()
    event.name = name
    event.begin = begin
    if end is not None:
        event.end = end
    else:
        event.duration = duration

    if location is not None:
        event.location = location

    if description is not None:
        event.description = description
    return event


def test_calendar():
    cal = Calendar()
    e = create_event('My wonderful event', datetime.datetime(2021, 10, 14, 11, 0), location='Chełmońskiego 9, Warszawa')
    cal.events.add(e)
    for event in read_from_excel():
        cal.events.add(event)

    print(cal)

    with open('test.ics', 'w', encoding='utf-8') as f:
        f.write(str(cal))


def calendar_from_excel(filename):
    cal = Calendar()
    for event in read_from_excel(filename):
        cal.events.add(event)

    return cal


if __name__ == '__main__':
    test_calendar()
