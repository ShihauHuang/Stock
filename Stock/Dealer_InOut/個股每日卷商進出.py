import requests
import re
from bs4 import BeautifulSoup
from urllib.request import urlretrieve
import cv2 as cv
import pygsheets

Stock_ID = 2497

gc = pygsheets.authorize(service_file='Dealer-InOut-827f15aeec31.json') 
sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1T-VbovBG1pCGmM2wTQOR-VH-5BOtfNZLvTmqck-Qqi4/')
dealers = sheet.worksheet_by_title('券商分點進出')

payloads={
    "stk_code" : str(Stock_ID),
    "charset" : "UTF-8"
}

url = "https://www.tpex.org.tw/web/stock/aftertrading/broker_trading/download_ALLCSV.php"

res = requests.post(url , data=payloads).text
dealer_list = re.findall('"\d+",".+?".".+?",".+?",".+?"' , res)
#dealer_list = list(map(lambda x:x.replace('"' , '') , dealer_list))

result_list = []
dealer_self_list = []


now_dealer = ''
buy_count = 0
sell_count = 0
buy_price_sum = 0
sell_price_sum = 0

if dealer_list != [] :

    url = "https://www.tpex.org.tw/web/stock/aftertrading/broker_trading/brokerBS.php?l=zh-tw"
    get_name = requests.post(url , data=payloads).text
    stock_name = re.search(str(Stock_ID) + '(.+)?</td>' , get_name).group(1).split(';')[1]
    stock_name = str(Stock_ID) + '_' + stock_name

    if dealer_list[-1].strip() == '' :
        dealer_list.pop()

    result_list.append( [stock_name , '卷商' , '價格' , '買進股數' , '賣出股數'] )
    for i in dealer_list :
        tmp = list(map(lambda x:x.replace('"' , '').replace(' ' , '').replace(',' , '') , i.split('","')))
        result_list.append(tmp)

        if now_dealer == '' : # 第一次
            now_dealer = tmp[1]
            buy_count = buy_count + int(tmp[3])
            sell_count = sell_count + int(tmp[4])
            buy_price_sum = buy_price_sum + (float(tmp[2]) * int(tmp[3]))
            sell_price_sum = sell_price_sum + (float(tmp[2]) * int(tmp[4]))

        elif now_dealer == tmp[1] :
            buy_count = buy_count + int(tmp[3])
            sell_count = sell_count + int(tmp[4])
            buy_price_sum = buy_price_sum + (float(tmp[2]) * int(tmp[3]))
            sell_price_sum = sell_price_sum + (float(tmp[2]) * int(tmp[4]))

        else : 
            if buy_count != 0 :
                buy_price_avg = round(buy_price_sum / buy_count , 2)
            else :
                buy_price_avg = 0 
            if sell_count != 0 :
                sell_price_avg = round(sell_price_sum / sell_count , 2)
            else :
                sell_price_avg = 0
            buy_count = str(round(buy_count/1000 , 1)).split('.')[0]
            sell_count = str(round(sell_count/1000 , 1)).split('.')[0]
            overbuy = str(int(buy_count) - int(sell_count))

            dealer_self_list.append( [now_dealer , overbuy , 'tmp' , buy_count , buy_price_avg , sell_count , sell_price_avg] )

            now_dealer = tmp[1]
            buy_count = int(tmp[3])
            sell_count = int(tmp[4])
            buy_price_sum = float(tmp[2]) * int(tmp[3])
            sell_price_sum = float(tmp[2]) * int(tmp[4])
        
        if dealer_list.index(i) == (len(dealer_list) - 1 ) : #如果是最後一個了
            if buy_count != 0 :
                buy_price_avg = round(buy_price_sum / buy_count , 2)
            else :
                buy_price_avg = 0 
            if sell_count != 0 :
                sell_price_avg = round(sell_price_sum / sell_count , 2)
            else :
                sell_price_avg = 0
            buy_count = str(round(buy_count/1000 , 1)).split('.')[0]
            sell_count = str(round(sell_count/1000 , 1)).split('.')[0]
            overbuy = str(int(buy_count) - int(sell_count))

            dealer_self_list.append( [now_dealer , overbuy , 'tmp' , buy_count , buy_price_avg , sell_count , sell_price_avg] )

else :
    with requests.Session() as s :
        print(123)
        page = s.get('https://bsr.twse.com.tw/bshtm/bsMenu.aspx')
        print(456)
        soup = BeautifulSoup(page.content , 'lxml')

        Captcha_file = re.search('src="(.+?)"' , str(soup.findAll('img')[1])).group(1)
        Captcha_url = 'https://bsr.twse.com.tw/bshtm/' + Captcha_file
        urlretrieve(Captcha_url , 'tmp/Captcha.png')

        cv.imshow('input image' , cv.imread('tmp/Captcha.png')) 
        cv.waitKey(0)
        Captcha = str(input('驗證碼 : '))

        payload = {
            "__EVENTTARGET" : "" ,
            "__EVENTARGUMENT" : "" ,
            "__LASTFOCUS" : "" ,
            "__VIEWSTATE" : soup.select_one("#__VIEWSTATE")["value"] , 
            "__VIEWSTATEGENERATOR" : soup.select_one("#__VIEWSTATEGENERATOR")["value"] , 
            "__EVENTVALIDATION" : soup.select_one("#__EVENTVALIDATION")["value"] , 
            "RadioButton_Normal" : "RadioButton_Normal" ,
            "TextBox_Stkno" : str(Stock_ID) ,
            "CaptchaControl1" : Captcha ,
            "btnOK" : "查詢"
        }

        post_data = s.post('https://bsr.twse.com.tw/bshtm/bsMenu.aspx' , data =payload)
        
        get_name = s.get('https://bsr.twse.com.tw/bshtm/bsContent.aspx?v=t').text
        stock_name = re.search(str(Stock_ID) + '(.+)?</td>' , get_name).group(1).split(';')[1]
        stock_name = str(Stock_ID) + '_' + stock_name

        res = s.get('https://bsr.twse.com.tw/bshtm/bsContent.aspx').text

        dealer_list = res.replace('\n',',,').split(',,')

        
        result_list.append( [stock_name , '卷商' , '價格' , '買進股數' , '賣出股數'] )

        if dealer_list[-1].strip() == '' :
            dealer_list.pop()

        for i in dealer_list : 
            if i == "" or i[0].isnumeric() == False : # 過濾非必要資訊
                continue

            result_list.append( i.split(',') )

            if now_dealer == '' : # 第一次
                now_dealer = i.split(',')[1]
                buy_count = buy_count + int(i.split(',')[3])
                sell_count = sell_count + int(i.split(',')[4])
                buy_price_sum = buy_price_sum + (float(i.split(',')[2]) * int(i.split(',')[3]))
                sell_price_sum = sell_price_sum + (float(i.split(',')[2]) * int(i.split(',')[4]))

            elif now_dealer == i.split(',')[1] :
                buy_count = buy_count + int(i.split(',')[3])
                sell_count = sell_count + int(i.split(',')[4])
                buy_price_sum = buy_price_sum + (float(i.split(',')[2]) * int(i.split(',')[3]))
                sell_price_sum = sell_price_sum + (float(i.split(',')[2]) * int(i.split(',')[4]))

            else : 
                if buy_count != 0 :
                    buy_price_avg = round(buy_price_sum / buy_count , 2)
                else :
                    buy_price_avg = 0 
                if sell_count != 0 :
                    sell_price_avg = round(sell_price_sum / sell_count , 2)
                else :
                    sell_price_avg = 0
                buy_count = str(round(buy_count/1000 , 1)).split('.')[0]
                sell_count = str(round(sell_count/1000 , 1)).split('.')[0]
                overbuy = str(int(buy_count) - int(sell_count))

                dealer_self_list.append( [now_dealer , overbuy , 'tmp' , buy_count , buy_price_avg , sell_count , sell_price_avg] )

                now_dealer = i.split(',')[1]
                buy_count = int(i.split(',')[3])
                sell_count = int(i.split(',')[4])
                buy_price_sum = float(i.split(',')[2]) * int(i.split(',')[3])
                sell_price_sum = float(i.split(',')[2]) * int(i.split(',')[4])
            
            if dealer_list.index(i) == (len(dealer_list) - 1 ) : #如果是最後一個了
                if buy_count != 0 :
                    buy_price_avg = round(buy_price_sum / buy_count , 2)
                else :
                    buy_price_avg = 0 
                if sell_count != 0 :
                    sell_price_avg = round(sell_price_sum / sell_count , 2)
                else :
                    sell_price_avg = 0
                buy_count = str(round(buy_count/1000 , 1)).split('.')[0]
                sell_count = str(round(sell_count/1000 , 1)).split('.')[0]
                overbuy = str(int(buy_count) - int(sell_count))

                dealer_self_list.append( [now_dealer , overbuy , 'tmp' , buy_count , buy_price_avg , sell_count , sell_price_avg] )


dealers.clear(start='R' , end = 'V')
dealers.clear("A4:N" )
dealers.update_values(crange = "R1"  , values=result_list) # 

dealer_self_list = list(filter(lambda x : (int(x[1]) != 0 or int(x[3]) != 0) == True , dealer_self_list))
volume = 0
for v in range (len(dealer_self_list)) : 
    if int(dealer_self_list[v][3]) > 0 :
        volume = volume + int(dealer_self_list[v][3])
for v2 in range (len(dealer_self_list)) : 
    dealer_self_list[v2][2] = str(round(int(dealer_self_list[v2][1]) * 100 / volume , 2)) + '%'

buy_sort = sorted(dealer_self_list , key = lambda a:int(a[1]) , reverse=True)
buy_sort = list(filter(lambda x : int(x[1]) > 0 , buy_sort))
sell_sort = sorted(dealer_self_list , key = lambda a:int(a[1]))
sell_sort = list(filter(lambda x : int(x[1]) < 0 , sell_sort))

dealers.update_values(crange = "A4"  , values=buy_sort) # 寫入數據
dealers.update_values(crange = "H4"  , values=sell_sort) # 寫入數據
