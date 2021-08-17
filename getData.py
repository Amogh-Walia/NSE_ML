import matplotlib.pyplot as plt
from datetime import date, datetime
from sqlalchemy import create_engine
import pymysql
import pandas as pd
import pandas as pd
import xlwings as xw
from xlwings.constants import DeleteShiftDirection






today = date.today()
day = (today.strftime("%b_%d")).lower()

def GetData(file):
    put_table = 'put'+day
    call_table = 'call'+day

    sqlEngine = create_engine('mysql+mysqlconnector://root:15w60ps@localhost:3306/test', echo=False)

    dbConnection    = sqlEngine.connect()

    frame           = pd.read_sql("SELECT   c.Time,sum(p.changeinOpenInterest)-sum(c.changeinOpenInterest) as sigmaChangeInOI  FROM  "+call_table+" c INNER JOIN  "+put_table+" p  ON c.strikePrice = p.strikePrice and c.Time = p.Time group by c.Time", dbConnection);
    frame2          = pd.read_sql("SELECT   sum(p.openInterest) as put_oi,sum(c.openInterest) as call_oi  FROM  "+call_table+" c INNER JOIN "+put_table+" p  ON c.strikePrice = p.strikePrice and c.Time = p.Time  group by c.Time",dbConnection);
    frame3          = pd.read_sql("select (Sum(`Call Premium turnover`)-SUM(`Put Premium turnover` )) as delta_Premium,Time from "+put_table+" group by Time", dbConnection);
    frame4          = pd.read_sql("select underlyingValue from "+put_table+" where strikePrice = '12800'", dbConnection);
    time            = pd.read_sql("select Time from "+put_table, dbConnection) 
    time = time.iloc[-1,-1]

    wb = xw.Book(file)
    ws = wb.sheets["Sheet1"]
    ws.clear()

    ws.range("A1").options(index=False,header = True).value = frame
    ws.range("C1").options(index=False,header = True).value = frame2
    ws.range("E1").options(index=False,header = True).value = frame3
    ws.range("F1").options(index=False,header = True).value = frame4
    ws.range('F2').api.Delete(DeleteShiftDirection.xlShiftUp)
    for cell in ws.used_range:
        cell.number_format = 'General'
    '''
    for i in range(1,1000):
        st = 'F'+str(i)
        st = ws.range(st).value
        if st == None:
            ws.range(str(i)+':'+str(i)).api.Delete(DeleteShiftDirection.xlShiftUp)
            break
    '''
    wb.save()
    return time
