# seating_chart_creator
This creates a seating chart from an excel document for a variable number of guests with configurable sized tables.  
You can add pairs of people who want to sit together and pairs of people who don't want to sit together.
TABS INCLUDE:
    Seating Chart Creator.py
    Create Table Rosters from Seating Chart.py
    README.MD*

# Steps
1) Keep `Requests.xlsx` in this folder. It always has that name and two tabs:
   1) "Requests" — people who should sit together (Name, Sit with)
   2) "Antirequests" — people who should not sit together (Name, Do not sit with)
   Row 1 is the header. One pair per row. Names should match the attendee names in the export.
2) Download a new collection summary from Cheddar Up. Leave that file as-is. Do not rename its tabs.
3) Execute "Seating Chart Creator.py".
    a) F5 or the Green/Black Play Button (sideways triangle flag).
    b) Choose the new Cheddar Up export when the file window opens. Guest names are read from the 5th and 6th tabs only.
    c) If you have to redo it, right click on the right pane (Console) and click Quit.
4) If necessary, execute "Create Table Rosters From Seating Chart.py" which uses the output from step 3.

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