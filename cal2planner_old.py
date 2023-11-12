# Import the pypdf library
import pypdf

# Open the input PDF file in read mode
input_file = open("input.pdf", "rb")

# Create a PdfFileReader object
pdf_reader = pypdf.PdfReader(input_file)

# Get the number of pages in the PDF file
num_pages = len(pdf_reader.pages)

# Define the text to search for
text_to_search = "January Week 3 Friday, 20"

# Loop through the pages and find the page that contains the text
for i in range(num_pages):
    # Get the page object
    page = pdf_reader.pages[i]
    # Extract the text from the page
    text = page.extract_text()
    # Check if the text is in the page
    if text_to_search in text:
        # Print the page number
        print(f"Text found on page {i+1}")
        # Break the loop
        break

# Create a PdfFileWriter object
pdf_writer = pypdf.PdfWriter()

# Copy all the pages from the input file to the output file
for i in range(num_pages):
    pdf_writer.add_page(pdf_reader.pages[i])

# Get the page that contains the text
page_to_edit = pdf_writer.pages[i]

# Create a PdfFileMerger object
pdf_merger = pypdf.PdfMerger()

# Create a temporary PDF file with the text to add or update
temp_file = open("temp.pdf", "wb")
temp_writer = pypdf.PdfWriter()
temp_page = temp_writer.add_blank_page(width=page_to_edit.mediabox.width, height=page_to_edit.mediabox.height)
temp_page.insert_text("New text", x=100, y=100) # Change the text and the coordinates as needed
temp_writer.write(temp_file)
temp_file.close()

# Merge the temporary PDF file with the page to edit
pdf_merger.append(pypdf.PdfFileReader("temp.pdf"))
page_to_edit.mergePage(pdf_merger.getPage(0))

# Open the output PDF file in write mode
output_file = open("output.pdf", "wb")

# Write the output file
pdf_writer.write(output_file)

# Close the files
input_file.close()
output_file.close()