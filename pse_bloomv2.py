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
    
    if ':' in stockName:     
    	wiki = "https://www.bloomberg.com/quote/"+stockName
    else:
    	wiki = "https://www.bloomberg.com/quote/"+stockName+":PM"

    print(wiki)
    page = urllib.urlopen(wiki)
    soup = BeautifulSoup(page, "html.parser")
    soup.prettify()

    detailList = []
    mainList = []
    detailDict = {} 
    headerList = ['ticker','price','Open','Previous Close','Volume','priceDate','YTD Return','1 Yr Return','Day Range','52Wk Range','timestamp']

    details = soup.find("div", {"class":"detailed-quote show"})

    mainList.append(['ticker',str(soup.find("div", {"class":"ticker"}).get_text().strip())])
    mainList.append(['price',str(soup.find("div", {"class":"price"}).get_text().strip())])
    mainList.append(['priceDate',str(soup.find("div", {"class":"price-datetime"}).get_text().strip())])

    for child in details.find_all('div', {"class":"cell cell__mobile-basic"}):
        #print(child.prettify())
        detailList = []

        for child2 in child.find_all('div'):
      
            detailList.append(str(child2.get_text().strip()))
       
        mainList.append(detailList)

    detailDict = dict(mainList)    
    print('Writing File:'+stockName)
    #create a new worksheet if stockName does not exist
    if stockName not in str(listWorksheets):
        #Create worksheet
        wsheet = sheet.add_worksheet(stockName,1,1)        
        wsheet.insert_row(headerList, 1)	
        #because insert row adds another epty row after it
        wsheet.delete_row(2)
    else:
        #change wsheet to an existing sheet
        wsheet = sheet.worksheet(stockName)
    
    localtime = time.asctime( time.localtime(time.time()) )
    
    wsheet.append_row([detailDict['ticker'],detailDict['price'],detailDict['Open'],detailDict['Previous Close'],detailDict['Volume'],detailDict['priceDate'],detailDict['YTD Return'],detailDict['1 Yr Return'],detailDict['Day Range'],detailDict['52Wk Range'], localtime])        

    # Sleep so not to cause high network load on bloomberg servers
    #print('Sleeping')
    #time.sleep(1)
