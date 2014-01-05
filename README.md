excel-output-importer
=====================

## Intro

This repository is largely a public way to share an Excel template--with Visual Basic for Applications (VBA) scripts--for easily reading and refreshing SAS/STATA/R/etc output for table construction.

## Background

When working my dissertation around 2007, I had a relatively standard workflow among economists (and many social scientists) of:

1. writing data analysis code in SAS or STATA
2. exporting results out of SAS/STATA either through text-based log files or comma-separated values (csv) tables 
3. reading these tables into excel, and then parsing, rearranging and formating these values for presentation in Excel tables

This process can be very time-consuming across steps 2 and 3, especially while the types of analysis being run are evolving, or even when a given set of analyses are frequently being updated. I found myself being pretty frustrated, and imagined that there had to be a way to cut down on tedious labor, and save my time for more interesting/creative/productive tasks.

Writing code to automatically generate output--steps 1 and 2 above--is low effort once that code has been set up. The fundamental human effort involved in table making is determining the content, layout and formatting of a table. The missing element was an efficient way to get program output into Excel, and into the tables.

The `excel-output-importer.xlsm` file uses Visual Basic for Application scripts to automatically read in csv or text file output from SAS, STATA and R (see the "Usage" notes below for what types of tables will and will not work). Those output files are then imported so that they are centrally available, and references are created to their contents. Specifically, each table is given a named refererence (read references on "Range Names" such as [this one](http://spreadsheets.about.com/od/excel-2010-beginner-ecourse/ss/2010-12-01-excel-2010-basics-tutorial-ecourse-step6.htm)), along with its header and a column which contains a unique ID for that row. Those references allow users to write lookup functions (for example, see [these references](http://spreadsheets.about.com/od/excelslookupfunctions/)) in Excel to bring those contents.

Once the tables are laid out and the Excel lookup functions are set up to bring in the right values, then all of the tables can be fully updated with new analysis with one click of the "Run Import" button.

## Usage 

The Excel file excel-output-importer.xlsm has a table in the `control room` tab where the user supplies all information about tables that are to be imported. The required information is:

* File path
* File name (without extension)
* Sheet name -- required only if the table is to be imported from an Excel spreadsheet
* New name -- optional, in case the user wants to rename the table as it is imported, to make it more intuitive or easier/shorter to reference
* Length -- this is the calculated length of the file name (or new name). This value turns red if the name is long enough that the name of the table will be truncated when imported as its own tab.
* Ext -- this is where the file extension is specified
* Import? -- any non-blank character will indicate that the table on that line will be imported
* Replace? -- any non-blank character will indicate that the contents of the specified table will replace the contents of the current table if it exists
* Col to Range Name -- optional, the name of the column that contains unique references for each row. The default is that the left-most column will be assumed to have these unique references, and therefore will be the column that is named as its own range.
* Hide? -- any non-blank character will hide the table within the workbook. This is desirable to reduce clutter, so long as the user does not need to inspect the table's content.

The buttons on the right are:

* Run Import -- this imports and replaces the tables that are specified in the blue region, as indicated by the import and replace 
* Run Hide/Unhide -- this will hide and/or unhide sheets as indicated in the "Hide?" column
* Refresh Range Names -- this will reset range names for tables specified in the blue region. This is useful if the user needs to reset values in the "Col to Range Name" field.
* Go to Selected Sheet -- because this workbook can sometimes accumulate many sheets, this allows the user to click on the name of a sheet (in either the "Sheet name" or "New name" fields, as appropriate) and jump to that sheet
* Delete Broken Range Names -- mainly for tidiness, this macro deletes all range names with references which have been broken, typically occurring when older, unused sheets are deleted



## Some Troubleshooting



## Picking the Right Tool for the Job
