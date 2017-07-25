# DFGsheet
Python Dataframe from Google Sheet

```
df=gs.df("A1:T30") # convert the google sheet into a dataframe
```

# Prerequisite
  
  Configure Google Drive API and Service Accounts:
  ```
  https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html
  obtain a json file that stores your username, passwords, etc
  ```
  
# Install DFGsheet
```
  pip install dfgsheet
```
# Use DFGsheet

```
from dfgsheet.google import *

g=Gsheet("client_secret.json")   # open and connect google drive with the credentials obtained from prerequisite
gb=g["h5boss_v0.1-TR"]           # open the google workbook

gs=gb[4]                         # open the sheet by index, starting from 0
gs=gb['sheet_name']              # or open the sheet by name

df=gs.df()                       # convert the entire sheet to a dataframe
df1=gs.df('M40:T73')             # or convert the specified area into a dataframe
```

# Thanks
```
gspread https://github.com/burnash/gspread
```
