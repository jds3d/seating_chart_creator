# seating_chart_creator
This creates a seating chart from an excel document for a variable number of guests with configurable sized tables.  
You can add pairs of people who want to sit together and pairs of people who don't want to sit together.
TABS INCLUDE:
    Seating Chart Creator.py
    Create Table Rosters from Seating Chart.py
    README.MD*

# Steps
1) Create "Seating Chart Creator.xls" in the format of "Seating Chart Creator - Template.xls"
   1) It has to have the following sheets: 
      1) "Punchbowl_Event_Guest_List"
          a) Download from Punchbowl in csv format.
          b) Copy and paste the data from the csv into the 1st tab (Punchbowl_event_Guest_List) in Seating Chart Creator.xls
      2) "Requests"
          a) Edit this manually
      3) "Antirequests" 
          a) Edit this manually
          
          
GO TO TAB "Seatin Chart Creator.py" to start          
2) Execute "Seating Chart Creator.py" which will pick up that excel file.
    a) F5 or the Green/Black Play Button (sideways triangle flag).
    b) Note: if you have to redo it, right click on the right pane (Console) and click Quit.
3) If necessary, execute "Create Table Rosters From Seating Chart.py" which uses the output from step 2.

# Note:
xlrd library no longer supports xlsx as of March 2022, so make sure to use xls file formats.
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