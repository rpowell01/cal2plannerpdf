import datetime
import argparse
import os
import textwrap
import json

from collections import namedtuple


def events2_namedtuple(all_events):
    new_event_list = []
    event_tuple = namedtuple('event', 'Start Subject Duration')
    for each_event in all_events:
        each_event[0] = datetime.datetime.fromisoformat(each_event[0])
        new_event_tuple = event_tuple._make(each_event)
        new_event_list.append(new_event_tuple)
    return new_event_list


def process_each_days_events(all_events):
    event_list = []
    new_event_list = events2_namedtuple(all_events)
    event_count = 0
    previous_year = None
    previous_month = None
    previous_day = None
    while event_count < len(new_event_list):
        if event_count == 0 or (
            new_event_list[event_count].Start.year == previous_year
            and new_event_list[event_count].Start.month == previous_month
            and new_event_list[event_count].Start.day == previous_day
        ):
            wrapped_subject = textwrap.wrap(
                new_event_list[event_count].Subject, 50, break_long_words=True
            )
            subject = '\n'.join(wrapped_subject)
            event_end = (
                datetime.timedelta(minutes=new_event_list[event_count].Duration)
                + new_event_list[event_count].Start
            )
            event_end_str = event_end.strftime('%I:%M%p ')
            event_string = (
                new_event_list[event_count].Start.strftime('%I:%M%p - ')
                + event_end_str
                + '\n'
                + subject
                + '\n'
            )
            previous_year = new_event_list[event_count].Start.year
            previous_month = new_event_list[event_count].Start.month
            previous_day = new_event_list[event_count].Start.day
            event_list.append(event_string)
            event_count = event_count + 1
        else:
            json_singleDay = json.dumps(
                event_list, default=serialize_datetime, indent=2
            )
            print('Events for %i/%i/%i' % (previous_month, previous_day, previous_year))
            print(json_singleDay)
            event_list = []
            previous_year = new_event_list[event_count].Start.year
            previous_month = new_event_list[event_count].Start.month
            previous_day = new_event_list[event_count].Start.day

    return


# Define a custom function to serialize datetime objects
def serialize_datetime(obj):
    if isinstance(obj, datetime.datetime):
        return obj.isoformat()
    raise TypeError('Type not serializable')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Open json file of calendar \
                                                  events and process events for each \
                                                  day'
    )
    parser.add_argument(
        '--filename', required=True, help='filename to open and process'
    )

    args = parser.parse_args()
    SCRIPT_PATH = os.path.dirname(os.path.realpath(__file__))
    SCRIPT_NAME = os.path.split(os.path.realpath(__file__))[1]
    begin_date = datetime.datetime.today()
    begin_date_str = datetime.datetime.strftime(begin_date, '%m/%d/%Y')
    date_string = begin_date.strftime('%m-%d-%Y')

    print('Processing ' + SCRIPT_PATH + '\\' + args.filename + '...')
    with open(SCRIPT_PATH + '\\' + args.filename, 'r+') as f:
        json_from_file = json.load(f)
        f.close()
        # print(json_from_file)

    single_day_events = process_each_days_events(json_from_file)
