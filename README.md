# seating_chart_creator
This creates a seating chart from an excel document for a variable number of guests with configurable sized tables.  
You can add pairs of people who want to sit together and pairs of people who don't want to sit together.
TABS INCLUDE:
    Seating Chart Creator.py
    Create Table Rosters from Seating Chart.py
    cheddar_up.py (finds the newest Cheddar Up download)
    README.MD*

# Steps
1) Keep `Requests.xlsx` in this folder. It always has that name and two tabs:
   1) "Requests" — people who should sit together (Name, Sit with)
   2) "Antirequests" — people who should not sit together (Name, Do not sit with)
   Row 1 is the header. One pair per row. Names should match the attendee names in the export.
2) Execute "Seating Chart Creator.py".
    a) F5 or the Green/Black Play Button (sideways triangle flag). A window opens.
    b) Click "Download latest from Cheddar Up". Cheddar Up opens in your browser (sign in if it asks).
       Open the TVNJC collection, then Manage > Reports > Collection Summary. Leave the file in Downloads.
       The window picks it up when the download finishes, copies it to the "Cheddar Up exports" folder,
       saves it as the export to use, and offers to create the seating chart.
       If a newer export is already in Downloads, the button offers to use it right away.
       You can still click Browse… to pick an export yourself. Do not rename its tabs.
       Guest names are read from the 5th and 6th tabs only.
    c) Click "Create seating chart". If a group is larger than the limit, the window asks before seating them together. Then set the table numbers and the files are written to the output folder.
    d) Close the window when you are done. To start over, run the script again.
3) To rebuild table rosters from a seating chart you edited, use "Rosters from a seating chart" in that same window. It uses the guest details saved the last time a chart was created.

# Getting the Cheddar Up export
Cheddar Up has no public way for a program to download the Collection Summary spreadsheet.
Its API key is only on the Team plan, its developer docs are not public, and the endpoints
described publicly return payments and collections, not that spreadsheet. So the program never logs in to Cheddar Up and saves no password or token.
It opens Cheddar Up in your normal browser and watches the Downloads folder for a new
`*collection_summary*.xlsx`. Partial downloads (`.crdownload`, `.part`) are ignored until they finish.

Optional settings in `configuration.json` (not in git):
- `cheddar_up_url`: page to open, for example the TVNJC collection's Manage page. Default `https://my.cheddarup.com/`.
- `cheddar_up_collection_filter`: text the export's file name should contain. Default `TVNJC`.
- `downloads_folder`: where the browser saves files, if it is not the Windows Downloads folder.

To check from a command prompt: `python cheddar_up.py` prints the newest export in Downloads and its tabs 5 and 6.
`python cheddar_up.py --open --wait 600` opens Cheddar Up and waits up to 10 minutes for a new one.

# Tests
`python -m unittest discover -s tests` checks the download step with fake files and runs the whole
seating pipeline on a made-up export. The window test needs a display and is skipped without one.

The old Punchbowl template is `punchbowl_Seating Chart Creator - Template.xls`.

# Note:
The Cheddar Up export, Requests.xlsx, seating chart, and table rosters are xlsx files.
maddy--check .1 and get josh to change it?
right click "quit" if you have to start over
go to output for seating charts and table rosters
title of xls template to be for for creating table seating etc is: Seating Chart Creator.xls

If, after you get you've made your table roster you make changes to the tables...

    1.  run "Create Table rosters from seating chart.py"
    2.  making sure that the updated seating chart is in the first tab in FILE IN FOLDER 'OUTPUT'
   THAT GOT GENERATED FROM THE ORIGINAL before the changes were made.
    ( cut and paste he new seating chart into the first tab, AFTER  copying and renaming the original seating chart
     that was created from the original running of the program into new tab and name that 'original'
    3.  Don't forget to look behind all the windows on the screen for the window to allow you to chose which seating 
        chart to make the table rosters from. (THIS SEATING CHART TO SELECT WILL BE 