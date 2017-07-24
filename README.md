# Pysheet
Python Dataframe from Google Sheet

```
df=gs.df("A1:T30") # convert the google sheet into a dataframe
```

# Prerequisite
  
  Configure Google Drive API and Service Accounts:
  ```
  https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html
  ```
  
# Install Pysheet

# Use PySheet

```
from pysheet import *
bookname="h5boss_v0.1-TR" # your google workbook name
gcred="client_secret.json" # your google acount credentials, get it after prerequisite
g=Gsheet(gcred) # open and connect google drive
gb=g[bookname] # open the google workbook
gs=gb[4] # open the 4th sheet
df=gs.df() # convert the 4th sheet to a dataframe
df1=gs.df('M40:T73') # convert specified area into a dataframe
```

# Thanks
```
gspread https://github.com/burnash/gspread
```
