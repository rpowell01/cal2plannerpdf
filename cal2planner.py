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
    event = namedtuple("event", "Start Subject Duration")
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
            a.Subject = a.Subject + " (Recurring)"

        events.append(event(a.StartInStartTimeZone, a.Subject, a.Duration))
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
            event_end = datetime.timedelta(minutes=event.Duration) + event.Start
            event_end_str = event_end.strftime("%I:%M%p ")
            event_string = event.Start.strftime("%I:%M%p - ") \
                + event_end_str + " " \
                + event.Subject
            event_list.append(event_string)
    return event_list

def print_descr(annot,description):
    red = (1, 0, 0)
    blue = (0, 0, 1)
    gold = (1, 1, 0)
    green = (0, 1, 0)
    
    """Print a short description to the right of each annot rect."""
    annot.parent.insert_text(
        annot.rect.bl -2, "%s" % description, color=blue, fontsize=9, fontname="TiRo"
    )
    
def events2pdf(date2update, event_list):

    new_doc = False  # indicator if anything found at all

    red = (1, 0, 0)
    blue = (0, 0, 1)
    gold = (1, 1, 0)
    green = (0, 1, 0)

    day2process_week = str(day2process.isocalendar()[1])
    day2process_month = calendar.month_name[day2process.month]
    day2process_dayname = day2process.strftime("%A")
    day2process_daynumber = str(day2process.day)
    
    # Define the text to search for
    text_to_search = (
        day2process_month + "\n"
        + "Week " + day2process_week + "\n" 
        + day2process_dayname + ", " + day2process_daynumber
    )
    # print(
    # "Searching for '%s' in document '%s'"
    # % (text_to_search.replace("\n", " "), input_file.name)
    # )
    
    for page in input_file:  # scan through the pages
        locations = None
        locations = page.search_for(text_to_search)
        if locations:
            new_doc = True
            print("Adding Outlook events to '%s' on page %i" % (text_to_search.replace("\n", " "), page.number + 1))
            # for location in locations:
            #     page.add_highlight_annot(location)  # underline
            
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

                # page.insert_text(
                #     text_insert_location.bl + (135, 0), events2pdf
                # )
                
            events2pdf = "\n"
            # text_9am = []
            # text_10am = []
            # text_11am = []
            # text_12pm = []
            # text_01pm = []
            # text_02pm = []
            # text_03pm = []
            # text_04pm = []
            # text_05pm = []
            # text_06pm = []
            # text_07pm = []
            # text_08pm = []
            for event in event_list:
                # if event.find("08:") > -1 or event.find("09:") > -1 and event.find("AM - ") > -1:
                #     text_9am.append(event) 
                # if event.find("10:") > -1 and event.find("AM - ") > -1:
                #     text_10am.append(event) 
                # if event.find("11:") > -1 and event.find("AM - ") > -1:
                #     text_11am.append(event) 
                # if event.find("12:") > -1 and event.find("PM - ") > -1:
                #     text_12pm.append(event) 
                # if event.find("01:") > -1 and event.find("PM - ") > -1:
                #     text_01pm.append(event) 
                # if event.find("02:") > -1 and event.find("PM - ") > -1:
                #     text_02pm.append(event) 
                # if event.find("03:") > -1 and event.find("PM - ") > -1:
                #     text_03pm.append(event) 
                # if event.find("04:") > -1 and event.find("PM - ") > -1:
                #     text_04pm.append(event) 
                # if event.find("05:") > -1 and event.find("PM - ") > -1:
                #     text_05pm.append(event) 
                # if event.find("06:") > -1 and event.find("PM - ") > -1:
                #     text_06pm.append(event) 
                # if event.find("07:") > -1 and event.find("PM - ") > -1:
                #     text_07pm.append(event) 
                # if event.find("08:") > -1 and event.find("PM - ") > -1:
                #     text_08pm.append(event) 
                events2pdf = events2pdf + event + "\n"
            # page.insert_text(nine_am_location[0].tl + (25,6), text_9am, fontsize=9, fontname="TiRo")
            # page.insert_text(ten_am_location[0].tl + (25,6), text_10am, fontsize=9, fontname="TiRo")
            # page.insert_text(eleven_am_location[0].tl + (25.6), text_11am, fontsize=9, fontname="TiRo")
            # page.insert_text(twelve_pm_location[0].tl + (25.6), text_12pm, fontsize=9, fontname="TiRo")
            # page.insert_text(one_pm_location[0].tl + (25.6), text_01pm, fontsize=9, fontname="TiRo")
            # page.insert_text(two_pm_location[0].tl + (25.6), text_02pm, fontsize=9, fontname="TiRo")
            # page.insert_text(three_pm_location[0].tl + (25.6), text_03pm, fontsize=9, fontname="TiRo")
            # page.insert_text(four_pm_location[0].tl + (25.6), text_04pm, fontsize=9, fontname="TiRo")
            # page.insert_text(five_pm_location[0].tl + (25.6), text_05pm, fontsize=9, fontname="TiRo")
            # page.insert_text(six_pm_location[0].tl + (25.6), text_06pm, fontsize=9, fontname="TiRo")
            # page.insert_text(seven_pm_location[0].tl + (25.6), text_07pm, fontsize=9, fontname="TiRo")
            # page.insert_text(eight_pm_location[0].tl + (25.6), text_08pm, fontsize=9, fontname="TiRo")
            annot = page.add_freetext_annot(schedule_location[0] + (100, 160, 100, 160), events2pdf, fontsize=9, fontname="TiRo")
            info = annot.info
            info["title"] = "Outlook Events"
            annot.parent.insert_text(
                annot.rect.tl +(40,5), "%s" % "(Show Outlook Events)", color=blue, fontsize=9, fontname="TiRo"
             )

            page.add_highlight_annot(page.search_for("(Show Outlook Events)"))  # underline
            text_loc = page.search_for("(Show Outlook Events)")
            # annot.set_rect(text_loc[0]+ (-75, 5))
            # annot.set_popup(schedule_location[0] + (75, 170, 75, 170))
            annot.set_info(info)
            annot.update()
                
            return new_doc

# Open the input PDF file in read mode
input_file_name = "input.pdf"
input_file = fitz.open(input_file_name)

i = 0
while i <= 6:
    day2process = datetime.timedelta(days=i) + datetime.datetime.now()

    events = getCalendarEntries(day2process, 7)
    # for event in events:
    #     print(event)
    date_str=day2process.strftime("%m/%d/%Y")
    event_list = GetSingleDayEvents(events, date_str)
# today = datetime.datetime.strptime("11/14/2023", "%m/%d/%Y")

    if len(event_list) >= 1:
        new_doc = events2pdf(day2process, event_list)
    else:
        print("No outlook events for %s. skipping..." % date_str)
        
    i += 1


if new_doc:
    input_file.save("marked-" + input_file.name)