import requests
import re
from bs4 import BeautifulSoup
import pandas as pd
import pygsheets

Stock_ID = 2330

gc = pygsheets.authorize(service_file='Fund-37435451c7c1.json') 
sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1JJJWutdj7HNOnCR6DrxF6WjIT4IXyOBCd7gK3TMCbLM/')
Estimate_value = sheet.worksheet_by_title('價值估計')

value_list = []

url = "https://goodinfo.tw/StockInfo/StockDetail.asp?STOCK_ID=" + str(Stock_ID)
headers = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36"
}
res = requests.get(url , headers=headers)
res.encoding = 'utf-8'
Stock_Name = re.search("<title>(.+)"  , res.text).group(1).split(' ')[0].split(')')[1]
#目前淨值</nobr><br><nobr style='color:green'>20.32&nbsp;元<
Current_close = re.search("目前股價.+?>([\d.]+).+?元"  , res.text).group(1)
Current_net_value = re.search("目前淨值.+?>([\d.]+).+?元"  , res.text).group(1)

value_list.append([str(Stock_ID) , Stock_Name , '淨值' , Current_net_value , '收盤價' , Current_close])
value_list.append(['年度' , '最高價' , '最低價' , 'EPS (元)' , '股利合計' , "ROE% (EPS/NAV)"])

url = "https://goodinfo.tw/StockInfo/StockDividendPolicy.asp?STOCK_ID=" + str(Stock_ID)
headers = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36"
}

res = requests.get(url , headers=headers)
res.encoding = 'utf-8'

soup = BeautifulSoup(res.text , 'lxml')
data = soup.select_one('#divDetail')

dfs = pd.read_html(data.prettify())
df = dfs[0]
df.columns = df.columns.get_level_values(3)
for index, row in df.iterrows():
    try :
        ROE = str(round(float(row['EPS  (元)']) / float(Current_net_value) * 100 , 2)) + '%'

        value_list.append([str(int(row['股利  發放  年度'])) , row['最高'] , row['最低'] , row['EPS  (元)'] , row['股利  合計'] , ROE])
    except Exception as e:
        #print(e)
        continue

Estimate_value.clear('*')
Estimate_value.update_values(crange = "A1"  , values=value_list) # 在寫入數據'''