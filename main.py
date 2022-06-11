#Copyright (c) 2022 Efe Akaröz
import os
from concurrent.futures import thread
import pandas as pd
import json
from smartapi import SmartConnect, SmartWebSocket 
import time
from datetime import datetime
import datetime
import random

os.system("clear")
print("#Copyright (c) 2022 Efe Akaröz\n")
print("Titles for excel")
print("'symbol' and 'token'")
fileinp = input("\nInput file(xlsx):")


excel_file = fileinp
theList = []
jsonlist = []


df = pd.read_excel(excel_file)
dataJson = json.loads(df.to_json())

for d in dataJson["symbol"]:

    if len(jsonlist) == 0:
        youcan=True
    for j in jsonlist:
        symbolname = j["symbol"]
        try:
            symbolname.split(dataJson["symbol"][d])[1]
            youcan = False
            break
        except:
            youcan=True

    if youcan == True:
        jsonlist.insert(0,{"symbol":dataJson["symbol"][d],"token":dataJson["token"][d],"ex":dataJson["exch_seg"][d]})




df2 = pd.read_excel("angelscript.xlsx")
dataJson2 = json.loads(df2.to_json())



obj=SmartConnect(api_key="y53Jh9dj")
data = obj.generateSession("I57428","Izan@123")
refreshToken= data['data']['refreshToken']
feedToken=obj.getfeedToken()
userProfile= obj.getProfile(refreshToken)
data = [

                  
       ]
for j in jsonlist:
    exchange=j["ex"]
    try:
        LTP = obj.ltpData(exchange, j["symbol"], j["token"])
        thenum = int(jsonlist.index(j))
        os.system("clear")
        print("{}%".format(int((100*(thenum+1))/len(jsonlist))))

        data.insert(0,[LTP["data"]["tradingsymbol"], LTP["data"]["symboltoken"], LTP["data"]["exchange"],LTP["data"]["ltp"],LTP["data"]["close"],LTP["data"]["high"], LTP["data"]["low"],LTP["data"]["open"]])
    except:
        data.insert(0,["Module Error", "", "","","","", "","",""])

data.insert(0,["symbol","token","exchange","ltp","close","high","low","open"])



import openpyxl


xlsx = openpyxl.Workbook()


sheet = xlsx.active



for row in data:

    sheet.append(row)


xlsx.save('appending.xlsx')






