# Import the fitz library
import fitz

import win32com.client
import datetime
from collections import namedtuple
 
 
event = namedtuple("event", "Start Subject Duration")
 
 
def get_date(datestr):
    try:  # py3
        adate = datetime.datetime.fromtimestamp(datestr.Start.timestamp())
    except Exception:
        adate = datetime.datetime.fromtimestamp(int(datestr.Start))
    return adate
 
 
def getCalendarEntries(days=1, dateformat="%d/%m/%Y"):
    """
    Returns calender entries for days default is 1
    Returns list of events
    """
    Outlook = win32com.client.Dispatch("Outlook.Application")
    ns = Outlook.GetNamespace("MAPI")
    appointments = ns.GetDefaultFolder(9).Items
    appointments.Sort("[Start]")
    appointments.IncludeRecurrences = "True"
    today = datetime.datetime.today()
    begin = today.date().strftime(dateformat)
    tomorrow = datetime.timedelta(days=days) + today
    end = tomorrow.date().strftime(dateformat)
    appointments = appointments.Restrict(
        "[Start] >= '" + begin + "' AND [END] <= '" + end + "'")
    events = []
    for a in appointments:
        adate = get_date(a)
        events.append(event(adate, a.Subject, a.Duration))
    return events
 
 
if __name__ == "__main__":
    events = getCalendarEntries()

# Open the input PDF file in read mode
input_file_name = "input.pdf"
input_file = fitz.open(input_file_name)

# Define the text to search for
text_to_search = "January\nWeek 3\nFriday, 20"
text_to_add = "My calendar entry"

print("underlining words containing '%s' in document '%s'" % (text_to_search, input_file.name))

new_doc = False  # indicator if anything found at all

for page in input_file:  # scan through the pages
		locations = None
		locations = page.search_for(text_to_search)
		if locations:
			new_doc = True
			print("found '%s' on page %i" % (text_to_search, page.number + 1))
			for location in locations:
				page.add_highlight_annot(location)  # underline
    
			text_insert_location = page.search_for("2 PM")
			if text_insert_location:
				text_insert_location = fitz.Rect(text_insert_location[1])
				page.insert_text(text_insert_location.bl + (135,0),"Outlook Events:\nMy inserted text")



if new_doc:
    input_file.save("marked-" + input_file.name)

# # Loop through the pages and find the page that contains the text
# for i in range(input_file.page_count):
#     # Get the page object
#     page = input_file[i]
#     # Extract the text from the page
#     text = page.get_text()
#     # Check if the text is in the page
#     if text_to_search in text:
#         # Print the page number
#         print(f"Text found on page {i+1}")
#         pagenum_to_update = i
#         # Break the loop
#         break

# # Create a PdfWriter object
# output_file = fitz.open()

# # Copy all the pages from the input file to the output file
# for i in range(input_file.page_count):
#     output_file.insert_pdf(input_file, from_page=i, to_page=i)

# # Get the page that contains the text
# page_to_edit = output_file[pagenum_to_update]

# # Create a temporary PDF file with the text to add or update
# temp_file_name = "temp.pdf"
# temp_file = fitz.open()
# temp_page = temp_file.new_page(width=page_to_edit.rect.width, height=page_to_edit.rect.height)
# temp_page.insert_textbox(page_to_edit.rect, text_to_add, fontsize=12) # Change the text and the font size as needed
# temp_file.save(temp_file_name)
# temp_file.close()

# # Merge the temporary PDF file with the page to edit
# temp_file = fitz.open(temp_file_name)
# page_to_edit.show_pdf_page(page_to_edit.rect, temp_file, 0)

# # Open the output PDF file in write mode
# output_file_name = "output.pdf"
# output_file.save(output_file_name)

# # Close the files
# input_file.close()
# output_file.close()
# temp_file.close()