# Import the fitz library
import fitz

import win32com.client
import datetime, pytz, locale, calendar
from collections import namedtuple


event = namedtuple("event", "Start Subject Duration")

# set the timezone to US/Pacific
# timezone = pytz.timezone('US/Central')
# datetime.datetime.now(timezone)

# # get the current time in the default timezone
# now = datetime.datetime.now()
# print(now)

from collections import namedtuple

event = namedtuple("event", "Start Subject Duration")


def get_date(datestr):
    # tz = datetime.datetime.now().astimezone().tzinfo
    try:  # py3
        adate = datetime.datetime.fromtimestamp(datestr.Start.timestamp())
    except Exception:
        adate = datetime.datetime.fromtimestamp(int(datestr.Start))
    return adate


def getCalendarEntries(begin_date=datetime.datetime.today(), days=1):
    """
    Returns calender entries for x days default is 1
    Returns list of events
    """
    DATE_FORMAT = "%m/%d/%Y"
    Outlook = win32com.client.Dispatch("Outlook.Application")
    ns = Outlook.GetNamespace("MAPI")
    appointments = ns.GetDefaultFolder(9).Items
    appointments.Sort("[Start]")
    appointments.IncludeRecurrences = True
    # start_date = datetime.datetime.today()
    # begin = start_date.date().strftime(DATE_FORMAT)
    begin_string = begin_date.strftime(DATE_FORMAT)
    end = datetime.timedelta(days=days) + begin_date
    end_string = end.date().strftime(DATE_FORMAT)
    appointments = appointments.Restrict(
        "[Start] >= '" + begin_string + "' AND [End] <= '" + end_string + "'"
    )
    events = []
    for a in appointments:
        # adate = get_date(a)
        if a.IsRecurring:
            EventSubject = a.Subject + " (Recurring)"
            # EventSubject = EventSubject + " (Recurring)"
        events.append(event(a.StartInStartTimeZone, EventSubject, a.Duration))
    return events


def GetSingleDayEvents(all_events, date_str):
    date = datetime.datetime.strptime(date_str, "%m/%d/%Y")
    event_list = []
    for event in all_events:
        if (
            event.Start.year == date.year
            and event.Start.month == date.month
            and event.Start.day == date.day
        ):
            event_string = event.Start.strftime("%I:%M%p - ") \
                + event.Subject \
                + " - (" \
                + str(event.Duration) \
                + " Min)"
            event_list.append(event_string)
    return event_list

def print_descr(annot,description):
    """Print a short description to the right of each annot rect."""
    annot.parent.insert_text(
        annot.rect.bl -2, "%s" % description, color=blue, fontsize=9, fontname="TiRo"
    )

day2process = datetime.datetime.now()
events = getCalendarEntries(day2process, 7)
for event in events:
    print(event)

event_list = GetSingleDayEvents(events, date_str=day2process.strftime("%m/%d/%Y"))
# today = datetime.datetime.strptime("11/14/2023", "%m/%d/%Y")
day2process_week = str(day2process.isocalendar()[1])
day2process_month = calendar.month_name[day2process.month]
day2process_dayname = day2process.strftime("%A")
day2process_daynumber = str(day2process.day)

# Open the input PDF file in read mode
input_file_name = "input.pdf"
input_file = fitz.open(input_file_name)

# Define the text to search for
text_to_search = (
    day2process_month + "\n"
    + "Week " + day2process_week + "\n" 
    + day2process_dayname + ", " + day2process_daynumber
)

print(
    "Highlighting words containing '%s' in document '%s'"
    % (text_to_search, input_file.name)
)

new_doc = False  # indicator if anything found at all

red = (1, 0, 0)
blue = (0, 0, 1)
gold = (1, 1, 0)
green = (0, 1, 0)

for page in input_file:  # scan through the pages
    locations = None
    locations = page.search_for(text_to_search)
    if locations:
        new_doc = True
        print("found '%s' on page %i" % (text_to_search, page.number + 1))
        for location in locations:
            page.add_highlight_annot(location)  # underline
        
        displ = fitz.Rect(40, 0, 40, 0)    
        schedule_location = page.search_for("Schedule")
        nine_am_location = page.search_for("9 AM")
        ten_am_location = page.search_for("10 AM")
        eleven_am_location = page.search_for("11 AM")
        twelve_pm_location =  page.search_for("12 PM")
        one_pm_location =  page.search_for("1 PM")
        two_pm_location =  page.search_for("2 PM")
        three_pm_location =  page.search_for("3 PM")
        four_pm_location =  page.search_for("4 PM")
        five_pm_location =  page.search_for("5 PM")
        six_pm_location =  page.search_for("6 PM")
        seven_pm_location =  page.search_for("7 PM")
        eight_pm_location =  page.search_for("8 PM")

        if two_pm_location:
            text_insert_location = fitz.Rect(two_pm_location[1])
            events2pdf = "\n"
            for event in event_list:
                events2pdf = events2pdf + event + "\n"
            # page.insert_text(
            #     text_insert_location.bl + (135, 0), events2pdf
            # )
            annot = page.add_freetext_annot(schedule_location[0] + displ, events2pdf, fontsize=9, fontname="TiRo")
            print_descr(annot,description="(Show Outlook Events)")


if new_doc:
    input_file.save("marked-" + input_file.name)