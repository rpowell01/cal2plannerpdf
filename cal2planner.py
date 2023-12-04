import datetime
import calendar
import os
import textwrap

import tkinter as tk
import tkinter.font as tkFont
from tkinter import filedialog
from tkcalendar import DateEntry

# from tkinter import messagebox
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
    app.update_mb(message_text='Sending email with updated pdf attached to:' + to)
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


def events2notes(pdf_file, date2update, event_list):
    new_doc = False  # indicator if anything found at all
    date2update_week = str(date2update.isocalendar()[1])
    date2update_month = calendar.month_name[date2update.month]
    date2update_dayname = date2update.strftime('%A')
    date2update_daynumber = str(date2update.day)
    date2update_year = str(date2update.year)

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
        + '\n'
        + 'Notes'
    )

    for page in pdf_file:  # scan through the pages
        locations = None
        locations = page.search_for(text_to_search)
        if locations:
            new_doc = True
            app.update_mb(
                message_text="Adding Outlook events to Notes page '%s' on page %i"
                % (text_to_search.replace('\n', ' '), page.number + 1)
            )
            notes_location = page.search_for(date2update_year)
            box1 = fitz.Rect(notes_location[0] + (-8, 22, 135, 679))
            """
            We use a Shape object (something like a canvas) to output the text and the
            rectangles surrounding it for demonstration.
            """
            shape1 = page.new_shape()  # create Shape
            shape1.draw_rect(box1)  # draw rectangles
            shape1.finish(width=0.3, color=getColor('red'), fill=getColor('white'))
            shape1.commit()  # write all stuff to page /Contents

            # box2 = fitz.Rect(notes_location[0] + (-10, 30, 135, 23))
            box2 = box1 + (0, 0, 0, -649)
            """
            We use a Shape object (something like a canvas) to output the text
            and the rectangles surrounding it for demonstration.
            """
            shape2 = page.new_shape()  # create Shape
            shape2.draw_rect(box2)  # draw rectangles
            shape2.finish(
                width=0.3, color=getColor('red'), fill=getColor('LightSteelBlue')
            )
            # Now insert text in the rectangles. Font "Times" will be used
            # by default. A return code rc < 0 indicates insufficient space
            # (not checked here).
            if len(event_list) > 1:
                event_header = 'Outlook Event Notes'
            else:
                event_header = 'Note Topics'
            rc = shape2.insert_textbox(
                box2,
                event_header,
                color=getColor('blue'),
                align=1,
                fontsize=10.5,
            )
            shape2.commit()  # write all stuff to page /Contents

            # box3 = fitz.Rect(notes_location[0] + (-8, 46, 135, 679))
            box3 = box2 + (0, 20, 0, 649)
            """
            Use a Shape object (something like a canvas) to output the text
            and the rectangles surrounding it.
            """
            shape3 = page.new_shape()  # create Shape
            shape3.draw_rect(box3)  # draw rectangles
            shape3.finish(width=0.3, color=getColor('red'), fill=getColor('gainsboro'))

            # Now insert text in the rectangles. Font "Times" will be used
            # by default. A return code rc < 0 indicates insufficient space
            # (not checked here).
            event_count = 0
            for event in event_list:
                if len(event_list) > 0:
                    spacing = (679 - 46) / len(event_list)
                else:
                    spacing = 0
                event_location = box3 + (0, event_count * spacing, 0, 0)
                if len(event_list) > 1:
                    rc = shape3.insert_textbox(
                        event_location, event, color=getColor('blue'), fontsize=10.5
                    )
                    event_count += 1
                    if rc < 0:
                        app.update_mb(
                            message_text='Insufficient space in schedule '
                            'box to add event'
                        )
            shape3.commit()  # write all stuff to page /Content

            line_shape = page.new_shape()
            line_shape.draw_line(box1.tr + 2, box1.br + 2)
            line_shape.finish(color=getColor('black'), fill=getColor('black'))
            line_shape.commit()
    return new_doc


def distance(x, y):
    return abs(x - y)


def events2pdf(pdf_file, date2update, event_list):
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
            app.update_mb(
                message_text="Adding Outlook events to '%s' on page %i"
                % (text_to_search.replace('\n', ' '), page.number + 1)
            )

            schedule_location = page.search_for('Schedule')
            events2pdf = ''
            box1 = fitz.Rect(schedule_location[0] + (0, 15, 130, 465))
            """
            We use a Shape object (something like a canvas) to output the text and the
            rectangles surrounding it for demonstration.
            """
            shape1 = page.new_shape()  # create Shape
            shape1.draw_rect(box1)  # draw rectangles
            shape1.finish(width=0.3, color=getColor('red'), fill=getColor('white'))
            shape1.commit()  # write all stuff to page /Contents

            box2 = fitz.Rect(schedule_location[0] + (0, 15, 130, 23))
            """
            We use a Shape object (something like a canvas) to output the text
            and the rectangles surrounding it for demonstration.
            """
            shape2 = page.new_shape()  # create Shape
            shape2.draw_rect(box2)  # draw rectangles
            shape2.finish(
                width=0.3, color=getColor('red'), fill=getColor('LightSteelBlue')
            )
            # Now insert text in the rectangles. Font "Times" will be used
            # by default. A return code rc < 0 indicates insufficient space
            # (not checked here).
            rc = shape2.insert_textbox(
                box2, 'Outlook Events', color=getColor('blue'), align=1, fontsize=10.5
            )
            shape2.commit()  # write all stuff to page /Contents

            box3 = fitz.Rect(schedule_location[0] + (0, 35, 130, 465))
            """
            Use a Shape object (something like a canvas) to output the text
            and the rectangles surrounding it.
            """
            shape3 = page.new_shape()  # create Shape
            shape3.draw_rect(box3)  # draw rectangles
            shape3.finish(width=0.3, color=getColor('red'), fill=getColor('gainsboro'))

            # Now insert text in the rectangles. Font "Times" will be used
            # by default. A return code rc < 0 indicates insufficient space
            # (not checked here).
            for event in event_list:
                events2pdf = events2pdf + event + '\n'
            rc = shape3.insert_textbox(
                box3, events2pdf, color=getColor('blue'), fontsize=10.5
            )
            if rc < 0:
                app.update_mb(
                    message_text='Insufficient space in schedule box to add event list'
                )
            shape3.commit()  # write all stuff to page /Content
            links = page.get_links()
            notes_more_location = page.search_for('More')
            link_count = 0
            for link in links:
                # search through all page links looking for an "approximate match" for
                # existing "Notes | More" link location
                if (
                    distance(notes_more_location[0].x0, link['from'].x0) > 5
                    and distance(notes_more_location[0].x1, link['from'].x1) > 5
                    and distance(notes_more_location[0].y0, link['from'].y0) > 5
                    and distance(notes_more_location[0].y1, link['from'].y1) > 5
                ):
                    link_count += 1
                else:
                    # Found it!
                    new_link = links[link_count]
                    new_link['from'] = box3
                    page.insert_link(new_link)

            return new_doc


def start_processing(
    input_pdf_name, output_pdf_name, total_days_to_process, add_to_notes, mail_to
):
    pdf_file = fitz.open(input_pdf_name)
    i = 0
    while i <= total_days_to_process - 1:
        day2process = datetime.timedelta(days=i) + app.cal_start.get_date()
        date_str = day2process.strftime('%m/%d/%Y')
        events = get_calendar_entries(day2process, total_days_to_process)
        event_list = get_single_day_events(events, date_str)

        if len(event_list) >= 1:
            events_added = events2pdf(pdf_file, day2process, event_list)
            if add_to_notes:
                notes_added = events2notes(pdf_file, day2process, event_list)
        i += 1

    split_filename = os.path.split(input_pdf_name)

    if events_added or notes_added:
        pdf_file.save(split_filename[0] + '\\' + output_pdf_name)
        app.update_mb(
            message_text='Updated pdf saved: '
            + split_filename[0]
            + '\\'
            + output_pdf_name
        )
        pdf_file.close()

        if mail_to:
            send_mail(
                to=mail_to,
                subject=SCRIPT_NAME,
                body=output_pdf_name,
                attachment_name=split_filename[0] + '\\' + output_pdf_name,
            )
    app.update_mb(message_text='Done')


class App:
    def update_mb(self, message_text):
        lb_size = self.lb_messagebox.size()
        self.lb_messagebox.insert(lb_size + 1, message_text)
        self.lb_messagebox.yview_scroll(1, 'units')
        self.lb_messagebox.update()

    def btn_select_inputfile_command(self):
        self.update_mb(message_text='Gathering input file details...')
        self.lbl_input_filename['text'] = filedialog.askopenfilename(
            initialdir=SCRIPT_PATH, filetypes=[('PDF files', '*.pdf')]
        )
        enddate = self.cal_end.get_date()
        current_time = datetime.datetime.now()
        current_time = current_time.strftime('%I%M%p')
        dashed_date_str = enddate.strftime('%m%d%Y-') + current_time
        split_filename = os.path.split(self.lbl_input_filename['text'])
        output_pdf_filename = (
            split_filename[1].replace('.pdf', '') + '-' + dashed_date_str + '.pdf'
        )
        self.tb_output_filename.config(state='normal')
        self.tb_output_filename.insert(0, output_pdf_filename)
        self.tb_output_filename.update()
        self.btn_start.config(state='normal')

    def tb_days2process_changed(self, *args):
        if self.tb_output_filename['state'] == 'normal':
            self.update_mb(
                message_text='Updating output filename due to change in number of days'
                'to process value.'
            )
            enddate = self.cal_end.get_date()
            current_time = datetime.datetime.now()

            current_time = current_time.strftime('%I%M%p')
            dashed_date_str = enddate.strftime('%m%d%Y-') + current_time
            split_filename = os.path.split(self.lbl_input_filename['text'])
            output_pdf_filename = (
                split_filename[1].replace('.pdf', '') + '-' + dashed_date_str + '.pdf'
            )
            self.tb_output_filename.delete(0, tk.END)
            self.tb_output_filename.insert(0, output_pdf_filename)
            self.tb_output_filename.update()

    def cb_dailynotes_command(self):
        if cb_dailynotes_value.get():
            current_value = 'True'
        else:
            current_value = 'False'
        self.update_mb('Add events to Daily Notes is ' + current_value)

    def cb_email_command(self):
        if cb_email_value.get():
            current_value = 'True'
            self.tb_mailto.config(state='normal')
            self.tb_mailto.delete(0, tk.END)
            self.tb_mailto.insert(0, MAIL_TO)
            self.tb_mailto.update()
        else:
            current_value = 'False'
            self.tb_mailto.delete(0, tk.END)
            self.tb_mailto.insert(0, '')
            self.tb_mailto.update()
            self.tb_mailto.config(state='disabled')

        self.update_mb('Send output file as attachment is ' + current_value)

    def btn_start_command(self):
        self.update_mb(
            message_text='Start button pressed, adding outlook events to selected pdf'
        )
        start_processing(
            input_pdf_name=self.lbl_input_filename['text'],
            output_pdf_name=self.tb_output_filename.get(),
            total_days_to_process=(
                self.cal_end.get_date() - self.cal_start.get_date()
            ).days
            + 1,
            add_to_notes=cb_dailynotes_value.get(),
            mail_to=self.tb_mailto.get(),
        )

    def btn_quit_command(self):
        self.update_mb(message_text='Quit Button Pressed, Exiting...')
        exit()

    def __init__(self, root):
        # setting title
        root.title('Cal2Planner')
        # setting window size
        width = 591
        height = 412
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (
            width,
            height,
            (screenwidth - width) / 2,
            (screenheight - height) / 2,
        )
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        self.btn_select_inputfile = tk.Button(root)
        self.btn_select_inputfile['bg'] = '#f0f0f0'
        ft = tkFont.Font(family='Times', size=10)
        self.btn_select_inputfile['font'] = ft
        self.btn_select_inputfile['fg'] = '#000000'
        self.btn_select_inputfile['justify'] = 'center'
        self.btn_select_inputfile['text'] = 'Select Input File'
        self.btn_select_inputfile.place(x=20, y=30, width=111, height=30)
        self.btn_select_inputfile['command'] = self.btn_select_inputfile_command

        self.lbl_input_filename = tk.Label(root)
        self.lbl_input_filename['activebackground'] = '#f4f4f4'
        ft = tkFont.Font(family='Times', size=10)
        self.lbl_input_filename['font'] = ft
        self.lbl_input_filename['fg'] = '#333333'
        self.lbl_input_filename['justify'] = 'left'
        self.lbl_input_filename['text'] = 'None'
        self.lbl_input_filename['relief'] = 'sunken'
        self.lbl_input_filename.place(x=150, y=30, width=414, height=34)

        self.lbl_cal_start = tk.Label(root)
        ft = tkFont.Font(family='Times', size=10, weight='bold')
        self.lbl_cal_start['font'] = ft
        self.lbl_cal_start['fg'] = '#333333'
        self.lbl_cal_start['justify'] = 'left'
        self.lbl_cal_start['text'] = 'Start Date: '
        self.lbl_cal_start.place(x=15, y=80, width=167, height=30)

        self.frame_cal_start = tk.Frame(root)
        self.frame_cal_start.place(x=135, y=85)

        self.cal_start = DateEntry(
            self.frame_cal_start, selectmode='day', date_pattern='mm-dd-yyyy'
        )
        self.cal_start.grid(row=1, column=1, padx=15)
        self.cal_start.bind('<<DateEntrySelected>>', self.tb_days2process_changed)

        self.lbl_cal_end = tk.Label(root)
        ft = tkFont.Font(family='Times', size=10, weight='bold')
        self.lbl_cal_end['font'] = ft
        self.lbl_cal_end['fg'] = '#333333'
        self.lbl_cal_end['justify'] = 'left'
        self.lbl_cal_end['text'] = 'End Date: '
        self.lbl_cal_end.place(x=245, y=80, width=167, height=30)

        self.frame_cal_end = tk.Frame(root)
        self.frame_cal_end.place(x=360, y=85)

        initial_end_date = (
            datetime.timedelta(days=int(TOTAL_DAYS_TO_PROCESS))
            + datetime.datetime.now()
        )
        self.cal_end = DateEntry(
            self.frame_cal_end,
            selectmode='day',
            year=initial_end_date.year,
            month=initial_end_date.month,
            day=initial_end_date.day,
            date_pattern='mm-dd-yyyy',
        )
        self.cal_end.grid(row=1, column=1, padx=15)
        self.cal_end.bind('<<DateEntrySelected>>', self.tb_days2process_changed)

        self.lbl_output_filename = tk.Label(root)
        ft = tkFont.Font(family='Times', size=10, weight='bold')
        self.lbl_output_filename['font'] = ft
        self.lbl_output_filename['fg'] = '#333333'
        self.lbl_output_filename['justify'] = 'left'
        self.lbl_output_filename['text'] = 'Output File Name:'
        self.lbl_output_filename.place(x=20, y=135, width=110, height=30)

        self.tb_output_filename = tk.Entry(root)
        self.tb_output_filename['borderwidth'] = '1px'
        ft = tkFont.Font(family='Times', size=10)
        self.tb_output_filename['font'] = ft
        self.tb_output_filename['fg'] = '#333333'
        self.tb_output_filename['justify'] = 'left'
        self.tb_output_filename['text'] = 'None'
        self.tb_output_filename.place(x=150, y=135, width=411, height=30)
        self.tb_output_filename.config(state='disabled')

        self.cb_dailynotes = tk.Checkbutton(root, variable=cb_dailynotes_value)
        ft = tkFont.Font(family='Times', size=10)
        self.cb_dailynotes['font'] = ft
        self.cb_dailynotes['fg'] = '#333333'
        self.cb_dailynotes['justify'] = 'left'
        self.cb_dailynotes['text'] = 'Add Events to Daily Notes'
        self.cb_dailynotes.place(x=15, y=180, width=188, height=35)
        self.cb_dailynotes['offvalue'] = '0'
        self.cb_dailynotes['onvalue'] = '1'
        self.cb_dailynotes['command'] = self.cb_dailynotes_command

        self.cb_email = tk.Checkbutton(root, variable=cb_email_value)
        ft = tkFont.Font(family='Times', size=10)
        self.cb_email['font'] = ft
        self.cb_email['fg'] = '#333333'
        self.cb_email['justify'] = 'left'
        self.cb_email['text'] = 'Send Output File as Email Attachment'
        self.cb_email.place(x=27, y=220, width=224, height=35)
        self.cb_email['offvalue'] = False
        self.cb_email['onvalue'] = True
        self.cb_email['command'] = self.cb_email_command

        self.lbl_mailto = tk.Label(root)
        ft = tkFont.Font(family='Times', size=10, weight='bold')
        self.lbl_mailto['font'] = ft
        self.lbl_mailto['fg'] = '#333333'
        self.lbl_mailto['justify'] = 'left'
        self.lbl_mailto['text'] = 'Mail to Address:'
        self.lbl_mailto.place(x=20, y=260, width=110, height=30)

        self.tb_mailto = tk.Entry(root, textvariable=MAIL_TO)
        self.tb_mailto['borderwidth'] = '1px'
        ft = tkFont.Font(family='Times', size=10)
        self.tb_mailto['font'] = ft
        self.tb_mailto['fg'] = '#333333'
        self.tb_mailto['justify'] = 'left'
        self.tb_mailto['text'] = MAIL_TO
        self.tb_mailto.place(x=150, y=260, width=411, height=30)
        self.tb_mailto.config(state='disabled')

        self.btn_start = tk.Button(root)
        self.btn_start['bg'] = '#f0f0f0'
        ft = tkFont.Font(family='Times', size=10)
        self.btn_start['font'] = ft
        self.btn_start['fg'] = '#000000'
        self.btn_start['justify'] = 'center'
        self.btn_start['text'] = 'Start'
        self.btn_start.place(x=230, y=300, width=70, height=25)
        self.btn_start['command'] = self.btn_start_command
        self.btn_start.config(state='disabled')

        self.btn_quit = tk.Button(root)
        self.btn_quit['bg'] = '#f0f0f0'
        ft = tkFont.Font(family='Times', size=10)
        self.btn_quit['font'] = ft
        self.btn_quit['fg'] = '#000000'
        self.btn_quit['justify'] = 'center'
        self.btn_quit['text'] = 'Quit'
        self.btn_quit.place(x=320, y=300, width=70, height=25)
        self.btn_quit['command'] = self.btn_quit_command

        self.lb_messagebox = tk.Listbox(root)
        ft = tkFont.Font(family='Times', size=10)
        self.lb_messagebox['font'] = ft
        self.lb_messagebox['fg'] = '#333333'
        self.lb_messagebox['justify'] = 'left'
        self.lb_messagebox.insert(0, 'Proccessing Details...')
        self.lb_messagebox['relief'] = 'sunken'
        self.lb_messagebox.place(x=20, y=340, width=547, height=61)

        self.sb_messagebox = tk.Scrollbar(self.lb_messagebox, orient='vertical')
        self.sb_messagebox.pack(side='right', fill='y')
        # Attaching Listbox to Scrollbar
        # Since we need to have a vertical
        # scroll we use yscrollcommand
        self.lb_messagebox.config(yscrollcommand=self.sb_messagebox.set)

        # setting scrollbar command parameter
        # to listbox.yview method its yview because
        # we need to have a vertical view
        self.sb_messagebox.config(command=self.lb_messagebox.yview)


# begin main code processing
if __name__ == '__main__':
    SCRIPT_PATH = os.path.dirname(os.path.realpath(__file__))
    SCRIPT_NAME = os.path.split(os.path.realpath(__file__))[1]

    root = tk.Tk()
    cb_dailynotes_value = tk.IntVar()
    cb_email_value = tk.IntVar()
    TOTAL_DAYS_TO_PROCESS = tk.IntVar()
    TOTAL_DAYS_TO_PROCESS = 7
    MAIL_TO = tk.StringVar()
    MAIL_TO = 'Send-To-Kindle <rpowell0216_scribe@kindle.com>'

    app = App(root)
    root.mainloop()
