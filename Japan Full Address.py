# -*- coding: utf-8 -*-
"""
Created on Sun May 27 17:47:41 2018

@author: rayyu

Python 2.7
"""

import requests
import xlrd
import xlutils
from xlutils.copy import copy
import json

file_location = "C:/Users/rayyu/Desktop/Mid Term Project/Book1.xlsx"
workbook = xlrd.open_workbook(file_location)
sheet = workbook.sheet_by_index(0)
addressColumn = sheet.row_values(0)
addressColumnLength = len(addressColumn) + 1
addressColumnNumber = 1
addresses = sheet.col_values(addressColumnNumber,0,1002) # I limited to the first 1000 records for testing, so just remove this if you want to do that full sheet
wb = copy(workbook)
wSheet = wb.get_sheet(0)

wSheet.write(0, len(addressColumn), 'Latitude')
wSheet.write(0, len(addressColumn)+1, 'Longtitude')
wSheet.write(0, len(addressColumn)+2, 'Full Address')
print('Finished reading, converting to full address:')


addr = sheet.col_values(addressColumnNumber,0,1002) 


#print(resp_json_payload['results'][0]['geometry']['location'])
#print(resp_json_payload['results'][0]['formatted_address'])



# Iterate over each address in the excel file
for addr in range(1, len(addresses)):
    ad = addresses[addr]
    response = requests.get('https://maps.googleapis.com/maps/api/geocode/json?address=%s' % (ad) )
    resp_json_payload = response.json()
    #print(english)
    # Skip any empty cells
    if addresses[addr] == "":
        continue
   
    while True:
        try:
        # Translate the English to French
            lat = resp_json_payload['results'][0]['geometry']['location']['lat']
            lng = resp_json_payload['results'][0]['geometry']['location']['lng']
            add = resp_json_payload['results'][0]['formatted_address']
        except:
            pass
        txtLat = json.dumps(lat)
        txtLng = json.dumps(lng)
        txtAdd = json.dumps(add)

        #print(french.text)
        wSheet.write(addr,addressColumnLength-1,txtLat)
        wSheet.write(addr,addressColumnLength,txtLng)
        wSheet.write(addr,addressColumnLength+1,txtAdd)
        break
wb.save('C:/Users/rayyu/Desktop/Mid Term Project/AddressData.xls')
