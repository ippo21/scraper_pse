# 1. Read an excel file with all tickers
# 2. Loop through all tickers and get stock prices
# 3. Save daily data in excel sheet
#
#
# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import sys
import re
import requests
import urllib
import gspread
import time
from oauth2client.service_account import ServiceAccountCredentials


# Prepare conection to gsheets
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

#Open gsheet
sheet = client.open("pse_data1")

# Get all worksheets
listWorksheets = sheet.worksheets() 

listRecords = sheet.sheet1.col_values(1)
print(listRecords)

for stocks in listRecords:
    
    stockName = str(stocks.strip())
    
    if '/' in stockName:     
    	wiki = "http://quotes.wsj.com/PH/"+stockName
    else:
    	wiki = "http://quotes.wsj.com/PH/"+stockName

    print(wiki)
    page = urllib.urlopen(wiki)
    soup = BeautifulSoup(page, "html.parser")
    soup.prettify()
    
    ticker = str(soup.find("div", {"class":"tickerName"}).get_text().strip())
    price = str(soup.find("div", {"id":"quote_val"}).get_text().strip())
    priceDate = str(soup.find("div", {"id":"quote_dateTime"}).get_text().strip())
    
    print('Writing File:'+ticker)
    #create a new worksheet if stockName does not exist
    if stockName not in str(listWorksheets):
        #Create worksheet
        wsheet = sheet.add_worksheet(stockName,1,3)        
    else:
        #change wsheet to an existing sheet
        wsheet = sheet.worksheet(stockName)
    
    #append row
    wsheet.append_row([ticker,price,priceDate])
    
    # Sleep so not to cause high network load on bloomberg servers
    print('Sleeping')
    time.sleep(3)
