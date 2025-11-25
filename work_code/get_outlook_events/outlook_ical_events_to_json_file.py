import datetime
from datetime import timedelta
import arrow
import os
import json
import argparse

from ics import Calendar
import requests

from collections import namedtuple

# # library to access outlook client
# import win32com.client

from dateutil.tz import tzlocal

import icalendar
import recurring_ical_events


def events_from_ical(url, begin_date=datetime.datetime.today(), days=7):
    url = url
    event = namedtuple('event', 'Start Subject Duration')
    begin_date = arrow.get(begin_date)
    # end_date = datetime.timedelta(days=days) + begin_date
    end_date = begin_date.shift(days=days)
    print('Getting ICS Calendar, please wait...')
    cal_file = requests.get(url, timeout=30).text
    c = Calendar(cal_file)
    print('Calendar retrieved.')
    # appointments = list(c.timeline.included(begin_date, end_date))
    appointments = list(c.events)
    # print("appointment list length:" + str(len(appointments)))
    appt_list = []
    for a in appointments:
        recurring = False
        if a.begin >= begin_date and a.end <= end_date:
            for contentline in a.extra:
                if contentline.name == 'RECURRENCE-ID':
                    a.name = a.name + ' (Recurring)'
                    recurring = True
            duration_minutes = (a.duration.total_seconds() / 3600) * 60
            if not recurring:
                if not a.all_day:
                    appt_list.append(
                        event(a.begin.to(local), a.name, int(duration_minutes))
                    )
                else:
                    appt_list.append(event(a.begin, a.name, int(duration_minutes)))

    calendar = icalendar.Calendar.from_ical(cal_file)
    recurring_events = recurring_ical_events.of(calendar).between(
        begin_date.datetime, end_date.datetime
    )
    for revent in recurring_events:
        if revent.name == 'VEVENT':
            # print(component.get("name"))
            print(revent.get('summary'))
            # print(component.get("description"))
            # print(component.get("organizer"))
            # print(component.get("location"))
            print(revent.decoded('dtstart'))
            print(revent.decoded('dtend'))
            duration_minutes = (
                (revent.decoded('dtend') - revent.decoded('dtstart')).total_seconds()
            ) / 60
            appt_list.append(
                event(
                    arrow.get(revent.decoded('dtstart')),
                    revent.get('summary') + ' (Recurring)',
                    int(duration_minutes),
                )
            )
    appt_list = sorted(appt_list, key=lambda x: [x.Start, x.Subject])
    return appt_list


# Define a custom function to serialize datetime objects
def serialize_datetime(obj):
    if isinstance(obj, arrow.arrow.Arrow):
        return obj.for_json()
    raise TypeError('Type not serializable')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Get outlook calendar events \
                                                  starting from a given start \
                                                  date and number of days'
    )
    parser.add_argument(
        '--ical_url',
        help='Internet url of the shared calendar (ical) to parse events from',
        default=(
            'https://outlook.office365.com/owa/calendar/'
            '09c7ba1d987348ba8d2cfd7d9eb9283c@tierpoint.com/'
            '9d45cc57bdb149d28e3d06528405b6954114801392010184494/calendar.ics'
        ),
    )
    parser.add_argument(
        '--startdate',
        help=(
            'Date string in the form of mm/dd/yyyy to start getting '
            'calendar events from. Default = current date'
        ),
        default=datetime.datetime.strftime(datetime.datetime.today(), '%m/%d/%Y'),
    )
    parser.add_argument(
        '--days',
        type=int,
        help='additional number of days from startdate to process. \
                              Default = 7',
        default=7,
    )

    args = parser.parse_args()

    # This contains the local timezone
    local = tzlocal()
    SCRIPT_PATH = os.path.dirname(os.path.realpath(__file__))
    SCRIPT_NAME = os.path.split(os.path.realpath(__file__))[1]
    begin_date = datetime.datetime.strptime(args.startdate, '%m/%d/%Y')
    begin_date_str = datetime.datetime.strftime(begin_date, '%m/%d/%Y')
    days = args.days
    end_date = begin_date + timedelta(days=days)
    end_date_str = datetime.datetime.strftime(end_date, '%m-%d-%Y')
    # url = (
    #     "https://outlook.office365.com/owa/calendar/"
    #     "09c7ba1d987348ba8d2cfd7d9eb9283c@tierpoint.com/"
    #     "9d45cc57bdb149d28e3d06528405b6954114801392010184494/calendar.ics"
    # )
    # events = get_calendar_entries(begin_date=begin_date, days=days+1)
    # events = test(url=args.ical_url)
    events = events_from_ical(url=args.ical_url, begin_date=begin_date, days=days)
    jsonObject = json.dumps(events, default=serialize_datetime, indent=2)
    # jsonObject = json.dumps(events, indent=2)
    print('Calendar events starting from %s for %d days' % (begin_date_str, days))

    print(jsonObject)
    date_string = begin_date.strftime('%m-%d-%Y')
    with open(
        SCRIPT_PATH
        + '\\outlook_events_'
        + date_string
        + '_thru_'
        + end_date_str
        + '.json',
        'w',
    ) as f:
        f.write(jsonObject)
        f.close()
