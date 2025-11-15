import datetime
from datetime import timedelta
import os
import json
import argparse

from collections import namedtuple

# library to access outlook client
import win32com.client


def get_calendar_entries(begin_date=datetime.datetime.today(), days=1):
    """
    Returns calender entries for x days default is 1
    Returns list of events
    """
    event = namedtuple('event', 'Start Subject Duration')
    date_format = '%m/%d/%Y'
    outlook = win32com.client.Dispatch('Outlook.Application')
    ns = outlook.GetNamespace('MAPI')
    appointments = ns.GetDefaultFolder(9).Items
    appointments.Sort('[Start]')
    appointments.IncludeRecurrences = True
    begin_string = begin_date.strftime(date_format)
    end = datetime.timedelta(days=days) + begin_date
    end_string = end.strftime(date_format)
    appointments = appointments.Restrict(
        "[Start] >= '" + begin_string + "' AND [End] <= '" + end_string + "'"
    )
    appt_list = []
    for a in appointments:
        if a.IsRecurring:
            a.Subject = a.Subject + ' (Recurring)'
        appt_list.append(event(a.StartInStartTimeZone, a.Subject, a.Duration))
    return appt_list


# Define a custom function to serialize datetime objects
def serialize_datetime(obj):
    if isinstance(obj, datetime.datetime):
        return obj.isoformat()
    raise TypeError('Type not serializable')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Get outlook calendar events \
                                                  starting from a given start \
                                                  date and number of days'
    )
    parser.add_argument(
        '--startdate',
        help='Date string in the form of mm/dd/yyyy to start getting \
                              calendar events from. Default = current date',
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
    SCRIPT_PATH = os.path.dirname(os.path.realpath(__file__))
    SCRIPT_NAME = os.path.split(os.path.realpath(__file__))[1]
    begin_date = datetime.datetime.strptime(args.startdate, '%m/%d/%Y')
    begin_date_str = datetime.datetime.strftime(begin_date, '%m/%d/%Y')
    days = args.days
    end_date = begin_date + timedelta(days=days)
    end_date_str = datetime.datetime.strftime(end_date, '%m-%d-%Y')
    events = get_calendar_entries(begin_date=begin_date, days=days + 1)
    jsonObject = json.dumps(events, default=serialize_datetime, indent=2)
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
