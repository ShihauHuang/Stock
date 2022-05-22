import requests
import re
import webbrowser


amount = {
    0 : "1-999" ,
    1 : "1,000-5,000" ,
    2 : "5,001-10,000" ,
    3 : "10,001-15,000" ,
    4 : "15,001-20,000" ,
    5 : "20,001-30,000" ,
    6 : "30,001-40,000" ,
    7 : "40,001-50,000" ,
    8 : "50,001-100,000" , 
    9 : "100,001-200,000" ,
    10 : "200,001-400,000" ,
    11 : "400,001-600,000" ,
    12 : "600,001-800,000" ,
    13 : "800,001-1,000,000" ,
    14 : "1,000,001以上"
}

Stock_ID = 6477
main_amount = 11
spreat_amount = 8

#chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"
#webbrowser.get(chrome_path + ' %s').open_new_tab('https://www.cmoney.tw/finance/stockmainkline.aspx?s=' + str(Stock_ID))
#webbrowser.get(chrome_path + ' %s').open_new_tab('https://fubon-ebrokerdj.fbs.com.tw/z/zc/zco/zco_' + str(Stock_ID) + '.djhtm')

while True :
    try :
        headers ={
            "User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36" , 
        }
        url='https://www.tdcc.com.tw/smWeb/QryStockAjax.do?REQ_OPR=qrySelScaDates'
        date_list = requests.post(url , headers = headers).text
        date_list = date_list[2:-2].split('","') # ['20210122', '20210115', '20210108', '20201231'....]

        newest_date = date_list[0]
        second_date = date_list[1]

        url='https://www.tdcc.com.tw/smWeb/QryStockAjax.do'
        payload={
            'scaDates':newest_date,
            'scaDate':newest_date,
            'SqlMethod':'StockNo',
            'StockNo':str(Stock_ID),
            'radioStockNo':str(Stock_ID),
            'StockName':'',
            'REQ_OPR':'SELECT',
            'clkStockNo':str(Stock_ID),
            'clkStockName':''
        }
        payload2={
            'scaDates':second_date,
            'scaDate':second_date,
            'SqlMethod':'StockNo',
            'StockNo':str(Stock_ID),
            'radioStockNo':str(Stock_ID),
            'StockName':'',
            'REQ_OPR':'SELECT',
            'clkStockNo':str(Stock_ID),
            'clkStockName':''
        }
        data = requests.post(url , headers = headers , data=payload).text
        data2 = requests.post(url , headers = headers , data=payload2).text
  
        percent_1 = re.findall('<td align="right">(.+)</td>\s+</tr>' , data)
        percent_2 = re.findall('<td align="right">(.+)</td>\s+</tr>' , data2)  

        newst_main_amount = 0
        second_main_amount = 0
        for p1 in range(main_amount , len(percent_1) - 1) :
            newst_main_amount = newst_main_amount + float(percent_1[p1])
        second_main_amount = 0
        for p2 in range(main_amount , len(percent_2) - 1) :
            second_main_amount = second_main_amount + float(percent_2[p2])

        try :
            how_many_up =  str(int((int(amount[main_amount].split('-')[0].replace(',' , ""))-1)/1000))
        except :
            how_many_up = "1000"

        if round(newst_main_amount , 2) > round(second_main_amount , 2) :
            is_up_or_down = "上升到"
        else :
            is_up_or_down = "下降到"
        print(re.search("證券名稱：(.+)</td>", data).group(1) + "_" + str(Stock_ID))
        print("持 " + how_many_up + " 張以上從 " + str(round(second_main_amount , 2)) + "% " + is_up_or_down + " " + str(round(newst_main_amount , 2)) + "%")


        newst_spreat_amount = 0
        second_spreat_amount = 0
        for p1 in range(spreat_amount , -1 , -1) :
            newst_spreat_amount = newst_spreat_amount + float(percent_1[p1])
        second_main_amount = 0
        for p2 in range(spreat_amount , -1 , -1) :
            second_spreat_amount = second_spreat_amount + float(percent_2[p2])

        how_many_down =  str(int(int(amount[spreat_amount].split('-')[1].replace(',' , ""))/1000))
        if round(newst_spreat_amount , 2) > round(second_spreat_amount , 2) :
            is_up_or_down = "上升到"
        else :
            is_up_or_down = "下降到"
        print("持 " + how_many_down + " 張以下從 " + str(round(second_spreat_amount , 2)) + "% " + is_up_or_down + " " + str(round(newst_spreat_amount , 2)) + "%")
        break

    except :
        continue

while True :
    try :
        headers ={
            "User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36" , 
        }
        url='https://www.tdcc.com.tw/smWeb/QryStockAjax.do?REQ_OPR=qrySelScaDates'
        date_list = requests.post(url , headers = headers).text
        date_list = date_list[2:-2].split('","') # ['20210122', '20210115', '20210108', '20201231'....]
        
        week_count = 20
        if len(date_list) >= week_count : 
            week = week_count
        else :
            week = len(date_list)

        for _date in range (0 , week) :
            while True :
                try :
                    url='https://www.tdcc.com.tw/smWeb/QryStockAjax.do'
                    payload={
                        'scaDates': date_list[_date] ,
                        'scaDate': date_list[_date] ,
                        'SqlMethod':'StockNo',
                        'StockNo':str(Stock_ID),
                        'radioStockNo':str(Stock_ID),
                        'StockName':'',
                        'REQ_OPR':'SELECT',
                        'clkStockNo':str(Stock_ID),
                        'clkStockName':''
                    }
                    data = requests.post(url , headers = headers , data=payload).text
                    percent = re.findall('<td align="right">(.+)</td>\s+</tr>' , data)

                    newst_main_amount = 0
                    for p1 in range(main_amount , len(percent) - 1) :
                        newst_main_amount = newst_main_amount + float(percent[p1])

                    newst_spreat_amount = 0
                    for p1 in range(spreat_amount , -1 , -1) :
                        newst_spreat_amount = newst_spreat_amount + float(percent[p1])  

                    print(date_list[_date] + " : " + str(round(newst_main_amount , 2)).split('.')[0] + '.' + str(round(newst_main_amount , 2)).split('.')[1].ljust(2 , "0") + '%' + "      " + str(round(newst_spreat_amount , 2)).split('.')[0] + '.' + str(round(newst_spreat_amount , 2)).split('.')[1].ljust(2 , "0") + '%')
                    break
                except :
                    continue


        break

    except :
        continue