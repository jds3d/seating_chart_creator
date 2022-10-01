# seating_chart_creator
This creates a seating chart from an excel document for a variable number of guests with configurable sized tables.  You can add pairs of people who want to sit together and pairs of people who don't want to sit together.

# Steps:
1) Create "Seating Chart Creator.xls" in the format of "Seating Chart Creator - Template.xls"
   1) It has to have the following sheets: 
      1) "Punchbowl_Event_Guest_List.1"
      2) "Requests"
      3) "Antirequests" 
2) Execute "Seating Chart Creator.py" which will pick up that excel file.
3) If necessary, execute "Create Table Rosters From Seating Chart.py" which uses the output from step 2.

# Note:
xlrd library no longer supports xlsx as of March 2022, so make sure to use xls file formats.
