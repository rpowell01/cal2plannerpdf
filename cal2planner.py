import datetime
import calendar
import os
import textwrap
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
# import json
import jsonpickle
from collections import namedtuple

# Import the fitz library for PyMuPDF
import fitz
from fitz.utils import getColor

# library to access outlook client
import win32com.client


def get_date(datestr):
    # tz = datetime.datetime.now().astimezone().tzinfo
    try:  # py3
        adate = datetime.datetime.fromtimestamp(datestr.Start.timestamp())
    except Exception:
        adate = datetime.datetime.fromtimestamp(int(datestr.Start))
    return adate


def send_mail(to, subject, body, attachment_name):
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    # mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

    # To attach a file to the email (optional):
    attachment = attachment_name
    mail.Attachments.Add(attachment)
    print("Sending email...")
    mail.Send()


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
            wrapped_subject = textwrap.wrap(event.Subject, 50,
                                            break_long_words=True)
            subject = "\n".join(wrapped_subject)
            event_end = datetime.timedelta(minutes=event.Duration) \
                + event.Start
            event_end_str = event_end.strftime("%I:%M%p ")
            event_string = event.Start.strftime("%I:%M%p - ") \
                + event_end_str + "\n" \
                + subject + "\n"
            event_list.append(event_string)

    if len(event_list) == 0:
        event_list.append("No events")
    return event_list
    
    
def links2json(filename):
    all_links = []
    # Open the input PDF file in read mode
    input_file = fitz.open(filename)
    output_file = os.path.splitext(filename)[0]+'.json'
    for page in input_file:  # scan through the pages
        page_links = page.links()
        all_links.append(page_links)
        
    json_object = jsonpickle.encode(all_links)
    # json_object = json.dumps(all_links, indent=4)
    with open(output_file, 'w', newline='') as myfile:
        myfile.write(json_object)
    return all_links
    
    
def json2links(filename):
    all_links = []
    with open(filename, 'rU', newline='') as json_file:
        json_str = json_file.read()
        all_links = jsonpickle.decode(json_str)
    json_file.close()
    return all_links
    
    
def links2pdf(input_file, all_links):
    page_counter = 0
    
    for page in input_file:  # scan through the pages
        for link in all_links[page_counter]:
            print(link)
        page_counter += 1
        

def events2pdf(date2update, event_list):

    new_doc = False  # indicator if anything found at all
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

    for page in input_file:  # scan through the pages
        locations = None
        locations = page.search_for(text_to_search)
        all_annots = page.annots()
        for annot in all_annots:
            page.delete_annot(annot)
        if locations:
            new_doc = True
            print("Adding Outlook events to '%s' on page %i" %
                  (text_to_search.replace("\n", " "), page.number + 1))

            schedule_location = page.search_for("Schedule")
            events2pdf = ""
            box1 = fitz.Rect(schedule_location[0] + (0, 15, 135, 480))
            """
            We use a Shape object (something like a canvas) to output the text
            and the rectangles surrounding it for demonstration.
            """
            shape1 = page.new_shape()  # create Shape
            shape1.draw_rect(box1)  # draw rectangles
            shape1.finish(width=0.3, color=getColor("red"),
                          fill=getColor("white"))
            shape1.commit()  # write all stuff to page /Contents

            box2 = fitz.Rect(schedule_location[0] + (0, 15, 135, 23))
            """
            We use a Shape object (something like a canvas) to output the text
            and the rectangles surrounding it for demonstration.
            """
            shape2 = page.new_shape()  # create Shape
            shape2.draw_rect(box2)  # draw rectangles
            shape2.finish(width=0.3, color=getColor("red"),
                          fill=getColor("LightSteelBlue"))
            # Now insert text in the rectangles. Font "Helvetica" will be used
            # by default. A return code rc < 0 indicates insufficient space
            # (not checked here).
            rc = shape2.insert_textbox(box2, "Outlook Events",
                                       color=getColor("blue"),
                                       align=1,
                                       fontsize=10.5)
            shape2.commit()  # write all stuff to page /Contents
            
            box3 = fitz.Rect(schedule_location[0] + (0, 35, 135, 480))
            """
            We use a Shape object (something like a canvas) to output the text
            and the rectangles surrounding it for demonstration.
            """
            shape3 = page.new_shape()  # create Shape
            shape3.draw_rect(box3)  # draw rectangles
            shape3.finish(width=0.3,
                          color=getColor("red"),
                          fill=getColor("gainsboro"))
            
            # Now insert text in the rectangles. Font "Helvetica" will be used
            # by default. A return code rc < 0 indicates insufficient space
            # (not checked here).
            for event in event_list:
                events2pdf = events2pdf + event + "\n"
            rc = shape3.insert_textbox(box3, events2pdf,
                                       color=getColor("blue"),
                                       fontsize=10.5)
            if rc < 0:
                print("Issuficiant space in schedule bax to add event list")
            shape3.commit()  # write all stuff to page /Content
            return new_doc

# generate links json from original planner file
# original_links = links2json("sn_a5x.breadcrumb.lined.default.ampm.sun.dailycal.2023.pdf")

# generate list of links from json file
# jsonfile_links = json2links("sn_a5x.breadcrumb.lined.default.ampm.sun.dailycal.2023.json")


total_days_to_process = 30
mail_to = 'Send-To-Kindle <rpowell0216_scribe@kindle.com>'
event = namedtuple("event", "Start Subject Duration")

root = tk.Tk()
root.withdraw()

scriptpath = os.path.dirname(os.path.realpath(__file__))
scriptname = os.path.split(os.path.realpath(__file__))[1]
input_file_name = filedialog.askopenfilename(initialdir=scriptpath,
                                             filetypes=[("PDF files",
                                                         "*.pdf")])
split_filename = os.path.split(input_file_name)
input_file = fitz.open(input_file_name)


i = 0
while i <= total_days_to_process -1:
    day2process = datetime.timedelta(days=i) + datetime.datetime.now()
    events = getCalendarEntries(day2process, total_days_to_process)
    date_str = day2process.strftime("%m/%d/%Y")
    event_list = GetSingleDayEvents(events, date_str)

    if len(event_list) >= 1:
        new_doc = events2pdf(day2process, event_list)
    else:
        print("No outlook events for %s. skipping..." % date_str)
    i += 1

# links2pdf(input_file,jsonfile_links)

dashed_date_str = day2process.strftime("%m%d%Y-%I%M%p")
outputfile = split_filename[1].replace(".pdf", "")+"-"+dashed_date_str+".pdf"
if new_doc:
    input_file.save(split_filename[0]+"\\"+outputfile)
    attachment_filename = split_filename[0]+"\\"+outputfile
    if tk.messagebox.askyesno(title=scriptname, message="Send updated pdf:\n\n" +
                              outputfile + "\n\nas email attachment?"):
        send_mail(to=mail_to,
                  subject=scriptname,
                  body=outputfile,
                  attachment_name=attachment_filename)