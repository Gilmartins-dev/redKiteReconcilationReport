# redKiteReconcilationReport
Checks that Gilmartins Total Value is less than Red Kites Total Value of Job


There is a video example of this program being run here https://gilmartins.sharepoint.com/sites/Gilmartins2/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FGilmartins2%2FShared%20Documents%2FGeneral%2FRecordings%2FNew%20channel%20meeting%2D20220104%5F120623%2DMeeting%20Recording%2Emp4&parent=%2Fsites%2FGilmartins2%2FShared%20Documents%2FGeneral%2FRecordings

Python program to reconcile Redkite’s Total Value with Gilmartins expected Total Value
This program identifies jobs from Redkite which need further inspection. 
They will have failed for one of two reasons. 
Gilmartins Job Value is greater than Redkites Job Value.
The status on Gilmartins Dynamics is not “invoiced”
We have added 10p to all Redkite total values to make sure that we are not throwing errors due to rounding.
This program is run on a spreadsheet supplied by Redkite, usually on a daily bases. At present it is managed by Tom Banner. He will usually send you the spreadsheet however you can request it from 
Robert Ilsley
Senior Project Manager
robert.ilsley@redkitehousing.org.uk 
07884207821
The code is held at https://github.com/Gilmartins-dev/redKiteReconcilationReport
It has comprehensive comments and a readme file.
Imports 
from os import times https://docs.python.org/3/library/time.html
this is used to format dates and times
import pandas as pd https://pandas.pydata.org/about/
This is used to read and write data
import hashlib https://docs.python.org/3/library/hashlib.html
This is used in the removeDuplicateLines() function on line 9
import time to manipulate time data
There is some useful boilerplate code which can be used in other routines
Line 24-36  getToken() requests the token needed to access the api this is stored in token on line 39
The token is then passed into getGilmartinsValue() on line 63 to get the value

To run this program all files must be in the same directory
Your input file will be named dailyreport.xlsx and is in excel format. Referenced on line 52
You must create a worksheet within the excel file, whatever name you give it should be placed in the variable sheetName on line 49
To prepare the worksheet
Copy the heading Job Number,	Total Value to the first two columns of the worksheet
In the Redkite spreadsheet copy the list of Job Number’s into notepad, then copy them from notepad into the dailyreport.xlsx
Do the same for the column marked Total Value
Save the spreadsheet and the python file (easily forgotten)
Run the python program from the command line : python redkiteReport.py
It should create three files in the same directory as the other files
sheetName_jobs_log_time this contains all jobs with the RedKite value next to the Gilmartins value for easy reference. If any jobs fail the criteria it will be marked with a fail in the right most column
sheetName_failed_list_dup_time this contains all the jobs that failed but will contain duplicates
sheetName_failed_log_time removes the duplicates. It is these that you send back to Tom
