import datetime
import calendar
import os
import textwrap

import tkinter
from tkinter import filedialog
from tkinter import messagebox
from collections import namedtuple

# Import the fitz library for PyMuPDF
import fitz
from fitz.utils import getColor

# library to access outlook client
import win32com.client


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
    print('Sending email...')
    mail.Send()


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
    end_string = end.date().strftime(date_format)
    appointments = appointments.Restrict(
        "[Start] >= '" + begin_string + "' AND [End] <= '" + end_string + "'"
    )
    appt_list = []
    for a in appointments:
        if a.IsRecurring:
            a.Subject = a.Subject + ' (Recurring)'
        appt_list.append(event(a.StartInStartTimeZone, a.Subject, a.Duration))
    return appt_list


def get_single_day_events(all_events, date_str):
    date = datetime.datetime.strptime(date_str, '%m/%d/%Y')
    event_list = []
    for event in all_events:
        if (
            event.Start.year == date.year
            and event.Start.month == date.month
            and event.Start.day == date.day
        ):
            wrapped_subject = textwrap.wrap(event.Subject, 50, break_long_words=True)
            subject = '\n'.join(wrapped_subject)
            event_end = datetime.timedelta(minutes=event.Duration) + event.Start
            event_end_str = event_end.strftime('%I:%M%p ')
            event_string = (
                event.Start.strftime('%I:%M%p - ')
                + event_end_str
                + '\n'
                + subject
                + '\n'
            )
            event_list.append(event_string)

    if len(event_list) == 0:
        event_list.append('No events')
    return event_list


def events2pdf(date2update, event_list):
    new_doc = False  # indicator if anything found at all
    date2update_week = str(date2update.isocalendar()[1])
    date2update_month = calendar.month_name[date2update.month]
    date2update_dayname = date2update.strftime('%A')
    date2update_daynumber = str(date2update.day)

    # Define the text to search for
    text_to_search = (
        date2update_month
        + '\n'
        + 'Week '
        + date2update_week
        + '\n'
        + date2update_dayname
        + ', '
        + date2update_daynumber
    )

    for page in pdf_file:  # scan through the pages
        locations = None
        locations = page.search_for(text_to_search)
        all_annots = page.annots()
        for annot in all_annots:
            page.delete_annot(annot)
        if locations:
            new_doc = True
            print(
                "Adding Outlook events to '%s' on page %i"
                % (text_to_search.replace('\n', ' '), page.number + 1)
            )

            schedule_location = page.search_for('Schedule')
            events2pdf = ''
            box1 = fitz.Rect(schedule_location[0] + (0, 15, 135, 480))
            """
            We use a Shape object (something like a canvas) to output the text and the
            rectangles surrounding it for demonstration.
            """
            shape1 = page.new_shape()  # create Shape
            shape1.draw_rect(box1)  # draw rectangles
            shape1.finish(width=0.3, color=getColor('red'), fill=getColor('white'))
            shape1.commit()  # write all stuff to page /Contents

            box2 = fitz.Rect(schedule_location[0] + (0, 15, 135, 23))
            """
            We use a Shape object (something like a canvas) to output the text
            and the rectangles surrounding it for demonstration.
            """
            shape2 = page.new_shape()  # create Shape
            shape2.draw_rect(box2)  # draw rectangles
            shape2.finish(
                width=0.3, color=getColor('red'), fill=getColor('LightSteelBlue')
            )
            # Now insert text in the rectangles. Font "Helvetica" will be used
            # by default. A return code rc < 0 indicates insufficient space
            # (not checked here).
            rc = shape2.insert_textbox(
                box2, 'Outlook Events', color=getColor('blue'), align=1, fontsize=10.5
            )
            shape2.commit()  # write all stuff to page /Contents

            box3 = fitz.Rect(schedule_location[0] + (0, 35, 135, 480))
            """
            Use a Shape object (something like a canvas) to output the text
            and the rectangles surrounding it.
            """
            shape3 = page.new_shape()  # create Shape
            shape3.draw_rect(box3)  # draw rectangles
            shape3.finish(width=0.3, color=getColor('red'), fill=getColor('gainsboro'))

            # Now insert text in the rectangles. Font "Helvetica" will be used
            # by default. A return code rc < 0 indicates insufficient space
            # (not checked here).
            for event in event_list:
                events2pdf = events2pdf + event + '\n'
            rc = shape3.insert_textbox(
                box3, events2pdf, color=getColor('blue'), fontsize=10.5
            )
            if rc < 0:
                print('Issuficiant space in schedule bax to add event list')
            shape3.commit()  # write all stuff to page /Content
            return new_doc


# begin main code processing
if __name__ == '__main__':
    scriptpath = os.path.dirname(os.path.realpath(__file__))
    scriptname = os.path.split(os.path.realpath(__file__))[1]
    total_days_to_process = 7
    mail_to = 'Send-To-Kindle <rpowell0216_scribe@kindle.com>'

    root = tkinter.Tk()
    root.withdraw()
    input_pdf_filename = filedialog.askopenfilename(
        initialdir=scriptpath, filetypes=[('PDF files', '*.pdf')]
    )

    pdf_file = fitz.open(input_pdf_filename)
    i = 0
    while i <= total_days_to_process - 1:
        day2process = datetime.timedelta(days=i) + datetime.datetime.now()
        date_str = day2process.strftime('%m/%d/%Y')
        events = get_calendar_entries(day2process, total_days_to_process)
        event_list = get_single_day_events(events, date_str)

        if len(event_list) >= 1:
            new_doc = events2pdf(day2process, event_list)
        i += 1

    dashed_date_str = day2process.strftime('%m%d%Y-%I%M%p')
    split_filename = os.path.split(input_pdf_filename)
    output_pdf_filename = (
        split_filename[1].replace('.pdf', '') + '-' + dashed_date_str + '.pdf'
    )
    if new_doc:
        print('Saving updated pdf: ' + split_filename[0] + '\\' + output_pdf_filename)
        pdf_file.save(split_filename[0] + '\\' + output_pdf_filename)
        pdf_file.close()
        attachment_filename = split_filename[0] + '\\' + output_pdf_filename
        if messagebox.askyesno(
            title=scriptname,
            message='Send updated pdf:\n\n'
            + output_pdf_filename
            + '\n\nas email attachment?',
        ):
            send_mail(
                to=mail_to,
                subject=scriptname,
                body=output_pdf_filename,
                attachment_name=attachment_filename,
            )
