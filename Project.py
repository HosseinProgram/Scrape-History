from datetime import datetime , timedelta
import requests
from copy import copy
import json
import os
from persiantools.jdatetime import JalaliDate
import openpyxl
from openpyxl.styles.borders import Border, Side
import threading
import time
with open('Config.json',encoding='utf-8') as json_file:
    Config=json.load(json_file)
with open("InsCodeDict.json",encoding='utf-8') as json_file:
    NamadDict=json.load(json_file)
try:
    os.makedirs('export')
except:
    pass
Symbols= Config['Symbols']
startdate = Config['startdate']
enddate = Config['enddate']
From = Config['From']
To = Config['To']
StepTime=Config['StepTime'] #Seconds

thin_border = Border(left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin'))


def MergeAndBorder(sheet,From, To):
    for i in range(From[0], To[0] + 1):
        for j in range(From[1], To[1] + 1):
            sheet.cell(row=i, column=j).border = thin_border
    sheet.merge_cells(start_row=From[0], start_column=From[1], end_row=To[0], end_column=To[1])

def BorderRange(sheet,From, To):
    for i in range(From[0], To[0] + 1):
        for j in range(From[1], To[1] + 1):
            sheet.cell(row=i, column=j).border = thin_border

def AllinMent(sheet,From, To):
    for i in range(From[0], To[0] + 1):
        for j in range(From[1], To[1] + 1):
            alignment_obj = copy(sheet.cell(row=i, column=j).alignment)
            alignment_obj.horizontal = 'center'
            alignment_obj.vertical = 'center'
            sheet.cell(row=i, column=j).alignment = alignment_obj


headers={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36 Edg/109.0.1518.61","Accept-Language":"en-US,en;q=0.9,fa;q=0.8","Connection": "keep-alive","Cookie": "ASP.NET_SessionId=kislljlalcplvzmn2q2ycni0"}



def GetSymbolHistory(symbol):
    
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    Start_Date=JalaliDate(startdate[0], startdate[1] , startdate[2]).to_gregorian()
    Start_Date=datetime(Start_Date.year, Start_Date.month, Start_Date.day)
    Start_Date=Start_Date.replace(hour=From[0],minute=From[1],second=From[2])

    End_Date = JalaliDate(enddate[0],enddate[1],enddate[2]).to_gregorian()
    End_Date = datetime(End_Date.year, End_Date.month, End_Date.day)

    index=NamadDict[symbol]
    datestartindex=int(str(Start_Date.year)+("0"+str(Start_Date.month))[-2:]+("0"+str(Start_Date.day))[-2:])
    dateendindex=int(str(End_Date.year)+("0"+str(End_Date.month))[-2:]+("0"+str(End_Date.day))[-2:])

    OrderRanking = [{},{},{},{},{}]

    while True:
        try:
            ActiveList=requests.get("http://www.tsetmc.com/tsev2/data/InstTradeHistory.aspx?i=%s&Top=999999&A=1"%index,timeout=5).text.split(";")
            break
        except:
            pass
            
    ActiveDates=list()
    for date in ActiveList:
        try: 
            ActiveDates.append(int(date.split("@")[0]))
        except: 
            pass
    ActiveDates.reverse()
    for i in range(0,len(ActiveDates)):
        date=ActiveDates[i]
        if date >= datestartindex:
            i1=i
            break
    for i in range(0,len(ActiveDates)):
        date=ActiveDates[i]
        if date >= dateendindex:
            i2=i
            break

    AllBourds=dict()

    j=i1
    while j <=i2:
        dateindex = ActiveDates[j]
        dateindex2=str(dateindex)
        print(symbol +" @ ", JalaliDate.to_jalali(int(dateindex2[0:4]) ,int(dateindex2[4:6]) ,int(dateindex2[6:8])).strftime("%Y/%m/%d"))
        try:
            BestLimits = requests.get('http://cdn.tsetmc.com/api/BestLimits/%s/%s'%(index,dateindex),headers=headers,timeout=5).json()["bestLimitsHistory"]
            CurrentClock = copy(Start_Date)
            EndClock = copy(Start_Date.replace(hour=To[0],minute=To[1],second=To[2]))
            Boards = list()
            while CurrentClock<=EndClock :
                CurrentHEven = int(str(CurrentClock.hour) + ("0"+str(CurrentClock.minute))[-2:] +("0"+str(CurrentClock.second))[-2:])

                for limit in BestLimits:
                    if CurrentHEven >= limit["hEven"] :
                        OrderRanking[limit["number"]-1]=limit
                    else:
                        break
                Boards.append([CurrentHEven,OrderRanking[0]["pMeDem"],OrderRanking[0]["pMeOf"]])
                CurrentClock = CurrentClock + timedelta(seconds=600)
            AllBourds[str(dateindex)]=Boards
            j+=1
        except Exception as e :
            pass


    BorderRange(sheet,[1,1], [len(list(AllBourds.keys()))+2,len(Boards)*2+1])
    AllinMent(sheet,[1,1], [len(list(AllBourds.keys()))+2,len(Boards)*2+1])
    sheet.cell(1,1).value="Time"
    sheet.cell(2,1).value="Date"
    for i in range(0,len(Boards)):
        boardstr=str(Boards[i][0])
        clock=boardstr[0:-4]+":"+boardstr[-4:-2]+":"+boardstr[-2:]
        sheet.cell(1,2*(i+1)).value=clock
        MergeAndBorder(sheet,[1,2*(i+1)], [1,2*(i+1)+1])
        sheet.cell(2,2*(i+1)).value="Buy"
        sheet.cell(2,2*(i+1)+1).value="Sale"


    for j in range(0,len(list(AllBourds.keys()))):
        dateindex = list(AllBourds.keys())[j]
        persiandate=JalaliDate.to_jalali(int(dateindex[0:4]) ,int(dateindex[4:6]) ,int(dateindex[6:8])).strftime("%Y/%m/%d")
        sheet.cell(j+3,1).value=persiandate
        Boards=AllBourds[dateindex]
        for i in range(0,len(Boards)):
            board=Boards[i]
            sheet.cell(j+3,2*(i+1)).value=board[1]
            sheet.cell(j+3,2*(i+1)+1).value=board[2]
    sheet.column_dimensions['A'].width=12
    sheet.freeze_panes ="A3"
    workbook.save("export/"+symbol+".xlsx")

k=0
for symbol in Symbols:
    k+=1
    print(k,symbol)
    threading.Thread(target=lambda: GetSymbolHistory(symbol)).start()
while threading.active_count() > 1:
    time.sleep(1)
print("======================================================================================")
print('The Program Is Done Successfully! \n\nPress any key to close the Program')
print("======================================================================================")
input()

