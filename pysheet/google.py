
# coding: utf-8

# In[454]:

import sys
import numpy as np
import pandas as pd
import gspread


# In[652]:

class ISheet:
    sheetname=""
    sheetid=1
    row=1
    col=1
    sheet=""
    def __init__(self, sheetname=None,sheetid=None):
        if sheetname:
            self.sheetname=sheetname
        if sheetid:
            self.sheetid=sheetid
    def _set_row(self, r):
        self.row=r
        
    def _set_col(self, c):
        self.col = c
        
    def _set_sheetid(self, sid):
        self.sheetid = sid
    
    def _sheetrange(self,sr):
        '''
            convert sheet range from label format to interger format, using sheet.get_int_addr
        '''
        srange=sr.split(':')
        topleft=srange[0]
        botright=srange[1]
        nrows=self.sheet.get_int_addr(botright)[0]-self.sheet.get_int_addr(topleft)[0]+1
        ncols=self.sheet.get_int_addr(botright)[1]-self.sheet.get_int_addr(topleft)[1]+1
        #print (nrows,ncols)
        return (nrows,ncols)
    
    def df(self, gsheet_range=None, header=None):
        '''
            return a dataframe given the label range
            if range not set, then return the max range
        '''
        gsheet=self.sheet # get the sheet object first
        if gsheet_range:
            try:
                range_obj=gsheet.range(gsheet_range)
                num_obj=len(range_obj)
                nrows,ncols=self._sheetrange(gsheet_range)
            except Exception as inst:
                print ('convert label format error')
                print (type(inst))
                print (inst)
                return
        else:
            # return the max area
            nrows,ncols=gsheet.row_count, gsheet.col_count
            num_obj=nrows*ncols
            range_obj=gsheet.range(1,1,nrows,ncols)
        try:
            allcellv=np.asarray([cell.value for cell in range_obj]).reshape(nrows,ncols)
            df = pd.DataFrame(allcellv)
        except Exception as inst:
            print ('convert sheet to dataframe error, "%s"'%gsheet_range)
            print (type(inst))
            print (inst)
            return
        return df


# In[653]:

class IBook:

    @property
    def bookname(self):
        return self._bookname
    @property
    def count(self):
        '''
            return the total number of sheets in a book
        '''
        return self._count
    @property
    def sheets(self):
        '''
            return a list that contains the sheets name available in the book
        '''
        return self._sheets
    @property
    def book(self):
        '''
            return a book object 
        '''
        return self._book
    def __init__(self, bookname=None):
        if bookname:
            self._bookname=bookname
    def _set_count(self, ct):
        self._count=ct
        
    def _set_sheets(self, sts):
        self._sheets = sts
        
    def _set_bookname(self, bkn):
        self._bookname = bkn
        
    #def Sheet(self, sheetid):
    def __getitem__(self, sheetid):
        '''
            return a sheet object based on the sheetid or sheetname
        '''
        isheet=ISheet()
        isheet.__init__()
        #self.sheetname=sheetname
        try:
            book=self.book
        except AttributeError:
            print ("no existing open book")
            return
        try:
            if not str(sheetid).isdigit():
                sheetid=self.sheets.index(sheetid)
            isheet.sheet=book.get_worksheet(sheetid)
            isheet.sheetname=isheet.sheet.title
            isheet.sheetid=sheetid
            return isheet
        except Exception as inst:
                print ('get google sheet error, "%s"'%sheetid)
                print (type(inst))
                print (inst)
                return      


# In[654]:

class Gsheet:
    def __init__(self, gcred):
        self._connect(gcred)
        
    def _connect(self,gcred):
        from oauth2client.service_account import ServiceAccountCredentials
        # use creds to create a client to interact with the Google Drive API
        scope = ['https://spreadsheets.google.com/feeds']
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(gcred, scope)
            self._client = gspread.authorize(creds)
            print ("google sheet is connected successfully")
        except Exception as inst:
            print ("connect to google sheet error using credential file, '%s'"%gcred)
            print (type(inst))
            print (inst)
            return

    #def Book(self, bookname):
    def __getitem__(self,bookname):
        '''
            return an ibook object based on the bookname
        '''
        ibook= IBook()
        ibook.__init__(bookname)
        try:
            client = self._client
        except AttributeError:
            print ("google sheet it not yet connected")
            return
        try:
            ibook._book=client.open(bookname)
            temp_sheets=ibook._book.worksheets()
            ibook._sheets=[str(itm).split('\'')[1] for itm in temp_sheets]
            ibook._count=len(ibook.sheets)
            return ibook
        except Exception as inst:
            print ("Open google book error, '%s'"%bookname)
            print (type(inst))
            print (inst)
            return
