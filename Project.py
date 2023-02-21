from datetime import datetime , timedelta
import requests
from copy import copy
import json
import os
from persiantools.jdatetime import JalaliDate
from pprintpp import pprint


symbol = "فولاد"
Desired_Time = [1395,7,10,11,27,16]




headers={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36 Edg/109.0.1518.61","Accept-Language":"en-US,en;q=0.9,fa;q=0.8","Connection": "keep-alive","Cookie": "ASP.NET_SessionId=kislljlalcplvzmn2q2ycni0"}
module_dir = os.path.dirname(__file__)
file_path = os.path.join(module_dir, 'InsCodeDict.json')

with open(file_path,encoding='utf-8') as json_file:
    NamadDict=json.load(json_file)


CurrentTime=JalaliDate(Desired_Time[0], Desired_Time[1] , Desired_Time[2]).to_gregorian()
CurrentTime=datetime(CurrentTime.year, CurrentTime.month, CurrentTime.day)
CurrentTime=CurrentTime.replace(hour=Desired_Time[3],minute=Desired_Time[4],second=Desired_Time[5])


index=NamadDict[symbol]
dateindex=str(CurrentTime.year)+("0"+str(CurrentTime.month))[-2:]+("0"+str(CurrentTime.day))[-2:]

OrderRanking = [{},{},{},{},{}]

BestLimits=requests.get('http://cdn.tsetmc.com/api/BestLimits/%s/%s'%(index,dateindex),headers=headers).json()["bestLimitsHistory"]
ClosingPrice=requests.get("http://cdn.tsetmc.com/api/ClosingPrice/GetClosingPriceHistory/%s/%s"%(index,dateindex),headers=headers).json()["closingPriceHistory"]
del ClosingPrice[-1]

CurrentHEven = int(str(CurrentTime.hour) + ("0"+str(CurrentTime.minute))[-2:] +("0"+str(CurrentTime.second))[-2:])

maxBuy=0
minSale = 0
for limit in BestLimits:
    if CurrentHEven >= limit["hEven"] :
        OrderRanking[limit["number"]-1]=limit
    else:
        break


pprint(OrderRanking[:5])


