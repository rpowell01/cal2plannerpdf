# cal2planner

## Overview

`cal2planner` is a Python script that retrieves calendar events / appointments from Outlook using the `win32com.client` and adds the retrieved events to a PDF based Planner file on the appropriate Daily Planner page for each of the event days retrieved.

The script creates a GUI front-end using the python tkinter library.  The GUI front-end allows you to:

* Select the PDF Planner file to update
* The begin and end date of calendar events to retrieve and embed in the PDF Planner File
* Define the output file name and wether or not it will be named based upon the end date selected
* Optionally add the events retrieved to the `Notes` page associated with the Planner Day
* Optionally send updated PDF Planner file as an Email Attachment and define the email recipient address

![](assets/20251125_105217_image.png)

## Planner PDF File - Notes

The script was created to specifically function against the PDF Planner files located at the following Link:
[pdf calendar download link](https://github.com/kudrykv/latex-yearly-planner/discussions/8)

Shown below are some sample images of the updated `Daily Planner` page and the optionally updated linked `Notes` page for the Daily planner page:

**Daily Planner**

![](assets/20251125_113249_image.png)

**Notes**

![](assets/20251125_113329_image.png)

## E-Reader Device - Notes

The have been tested against the planner files created for the `SuperNote ax5` and the planner files created for the `Remarkable2`.  The script will like work against the Planner files created for other devices, but this has not specifically been tested.

When I originally created the script, I had a Kindle Scribe, and the script worked relatively well. Hence the option to send the modified PDF Planner file to an email address like `Send-To-Kindle <username@kindle.com>` which I utilized the Amazon Send-To-Kindle functionality to view and update the modified PDF Planner file.  Unfortunately , the Kindle Scribe breaks the page-to-page embeded hyperlink within the PDF Planner file.  Not sure whether Amazon has ever addressed this problem as I no long use the Kindle Scribe.

The device that I am currently using is a `Boox Note Air 4c` and it works really well with the PDF Planner files.  Specifically the `SuperNote ax5` based Planner files work best for this device.  It should be noted that <u>when opening and editing the Planner files, you should specifically use the `Boox NeoReader`application and NOT the`Boox Notes` application</u>.  The Boox Notes application has the same issue as the Kindle Scribe, where it breaks the embedded page-to-page hyperlinks which serverely limits the functionality of the Planner Files.  The NeoReader application from my experience so far has all the same functionality as the Notes application, maybe more, and does not break the embedded hyperlinks.

### Synchronzing the updated PDF Planner File

Since the Boox Note Air 4c is an Android device at heart, I am able to utilize the `Syncthing` application installed on both the Windows machine and The Boox Note Air 4c tablet. The modifications to the planner PDF files are synchronized between devices nearly immediately after saving on either device.

Again, when using the Kindle Scribe, I used the "Send-To-Kindle" functionality to get the modified file from the Windows device to the Kindle Device, but it was not without its problems (see above).

## Prerequisites

Before running the script, make sure you have the following dependencies installed:

- Python 3.x
- PyMuPDF
- pypiwin32
- tkcalendar

## Usage

To use the script, follow these steps:

1. Optionally create and activate a python virtual environment.

   ```
   python -m venv .venv
   .\.venv\Scripts\activate
   ```
2. Install the required dependencies ).

   ```
   pip3 install -r ./requirements.txt
   ```
3. Run the script with the following command:

   ```shell
   python cal2planner.py
   ```
