import urllib.request
import datetime
import time
import os.path
import csv
from re import search
from requests import get
import pandas as pd
import json
import threading
import math

def Download_Data_by_date (_end_day) :

    now_year = _end_day.strftime("%Y%m%d")[0:4]
    now_month = _end_day.strftime("%Y%m%d")[4:6]

    # 以 _end_day 的月份找到當季的月初為哪一個月
    if now_month in ["01" , "02" , "03"] :
        _start_day = datetime.datetime.strptime(now_year + "0101" ,'%Y%m%d')
    elif now_month in ["04" , "05" , "06"] :
        _start_day = datetime.datetime.strptime(now_year + "0401" ,'%Y%m%d')
    elif now_month in ["07" , "08" , "09"] :
        _start_day = datetime.datetime.strptime(now_year + "0701" ,'%Y%m%d')
    elif now_month in ["10" , "11" , "12"] :
        _start_day = datetime.datetime.strptime(now_year + "1001" ,'%Y%m%d')
    
    try :
        print('開始下載投信資料...')
        twse_url = 'https://www.twse.com.tw/fund/TWT44U?response=csv&date=' #下載網址 上市
        tpex_url = 'https://www.tpex.org.tw/web/stock/3insti/sitc_trading/sitctr_download.php?l=zh-tw&t=D&type=buy&d=' #下載網址 上櫃
        
        #print("_start_day : " + str(_start_day) )
        #print("_end_day : " + str(_end_day) )
        
        totaldays = (_end_day - _start_day).days + 1
        for daynumber in range(totaldays):
            datestring = (_start_day + datetime.timedelta(days = daynumber)).date()
            Current_day = datestring.strftime("%Y%m%d")

            if not os.path.exists('Fund_Data/TWT44U_' + Current_day + '.csv') : # 如果不存在，則下載
                urllib.request.urlretrieve(twse_url + Current_day, "Fund_Data/TWT44U_" + Current_day + ".csv")
                print(Current_day)
            if not os.path.exists('Fund_Data/sit_' + Current_day + '.csv') : # 如果不存在，則下載
                urllib.request.urlretrieve(tpex_url + str(int(Current_day[0:4]) - 1911) + '/' + Current_day[4:6] + '/' + Current_day[6:8] , "Fund_Data/sit_" + Current_day + ".csv")
        print('下載投信資料完成')
        return _start_day
    except Exception as e :
        print('下載失敗 : ' + str(e))

def Get_All_Stock_from_Season_1th_to_Today(Today_timeformmat , Season_1st_month) :
    
    # 取得季初到當天投信買過的股票，放到 list ====================================================
    All_Stock = []
    Deal_Day = []

    totaldays = (Today_timeformmat - Season_1st_month).days + 1 # 取得今天與季初共有多少天
    for daynumber in range(totaldays): # 從季初日期開始跑迴圈到今天日期
        datestring = (Season_1st_month + datetime.timedelta(days = daynumber)).date()
        Current_day = datestring.strftime("%Y%m%d")
        Current_day_with_Slash = datestring.strftime("%Y/%m/%d")

        # 上市 ===============================================================================================
        with open ('Fund_Data/TWT44U_' + Current_day + '.csv' , newline='') as csvfile :
            _rows = csv.reader(csvfile, delimiter=',')
            rows = [_row for _row in _rows]
            if rows[0] != [] : # 上市與上櫃第一行目前來看都是 []，不是 [] 則算入一天
                for row in rows :
                    if row == [] :
                        continue
                    elif row[0] == ' ' and int(int(row[5].replace(',' , '')) / 1000) > 0:  # 若此判斷成立則為股票欄位，並同時為買超
                        All_Stock.append(row[2].strip() + '_' + search('(\w+)' , row[1]).group(1))
                
                Deal_Day.append(Current_day_with_Slash)

        # 上櫃 ===============================================================================================
        with open ('Fund_Data/sit_' + Current_day + '.csv' , newline='') as csvfile :
            _rows = csv.reader(csvfile, delimiter=',')
            rows = [_row for _row in _rows]   
            if rows[0] != [] : # 上市與上櫃第一行目前來看都是 []，不是 [] 則算入一天
                for row in rows :
                    if row == [] :
                        continue
                    elif row[0].isnumeric() and int(row[5].replace(',' , '')) > 0:  # 若此判斷成立則為股票欄位，並同時為買超
                        All_Stock.append(row[2].strip() + '_' + row[1].strip())
    
    #Deal_Day.reverse
    
    return list(set(All_Stock)) , Deal_Day

def Get_Total_Stock_Amounts(_stock_id) :
    url = "https://fubon-ebrokerdj.fbs.com.tw/z/zc/zcx/zcxD1.djjs?A=" + str(_stock_id)
    data = get(url)
    data.encoding = "big5"
    amounts = search("集保庫存.+?>([0-9,]+)" , data.text).group(1).replace("," , "")
    return amounts

def Date_to_Stamp(_date) :
    #_date need use the fommat YYYY-MM-DD
    timeString = str(_date) + " 08:00:00" # 時間格式為字串
    struct_time = time.strptime(timeString, "%Y-%m-%d %H:%M:%S") # 轉成時間元組
    time_stamp = int(time.mktime(struct_time)) # 轉成時間戳
    return str(time_stamp)

def Stamp_to_Date(_stamp) :

    struct_time = time.localtime(int(_stamp)) # 轉成時間元組
    #timeString = time.strftime("%Y-%m-%d %H:%M:%S", struct_time) # 轉成字串
    timeString = time.strftime("%Y/%m/%d", struct_time) # 轉成字串
    return timeString

def Get_Close_and_volume(_stock_Id , Today_timeformmat , Season_1st_month , Trade_Day_List) :

    """
    #https://ws.api.cnyes.com/ws/api/v1/charting/history?resolution=D&symbol=TWS:8358:STOCK&from=1609372800&to=1601510400
    #https://invest.cnyes.com/twstock/TWS/8358/history
    """
    url = "https://ws.api.cnyes.com/ws/api/v1/charting/history?resolution=D&symbol=TWS:" + _stock_Id + ":STOCK&from=" + Date_to_Stamp(Today_timeformmat.strftime('%Y-%m-%d')) + "&to=" + Date_to_Stamp(Season_1st_month.strftime('%Y-%m-%d'))
    #print(url)
    tmp_data = get(url)
    data = json.loads(tmp_data.text)

    Self_Trade_Day = list(map(lambda x : Stamp_to_Date(x) , data['data']['t']))
    Close_List = data['data']['c']
    Volumn_List = list(map(lambda x : int(round(x , 0)) , data['data']['v']))

    Self_Trade_Day.reverse()
    Close_List.reverse()
    Volumn_List.reverse()

    if len(Trade_Day_List) != len(Self_Trade_Day) : # 部分因為減資等等原因會暫停交易幾天
        for _check_start in Trade_Day_List : 
            if _check_start not in Self_Trade_Day :
                Close_List.insert(Trade_Day_List.index(_check_start) , Close_List[Trade_Day_List.index(_check_start) - 1])
                Volumn_List.insert(Trade_Day_List.index(_check_start) , 0)
    
    Close_List.reverse()
    Volumn_List.reverse()

    update_volumn_list = [] # 網站上資訊有誤
    total = int(Get_Total_Stock_Amounts(_stock_Id))

    for v in Volumn_List :
        if v > total : 
            update_volumn_list.append( int(round(v/1000 , 0)) )
        else :
            update_volumn_list.append(v)

    return Close_List , update_volumn_list

def Get_fund_buy_info(All_Stock , Today_timeformmat , Season_1st_month , Trade_Day_List) :

    Huge_list = []
    Error_List = []
    Today_overbuy_Stock_list = []

    def do_important_job (num) :

        interval = math.floor(len(All_Stock) / 50) # 無條件捨去
        if num == 50 :
            end = len(All_Stock)
        else :
            end = num * interval
        start = interval * num - interval

        for _Stock in range (start , end) : # 6285 啟碁 無解，少部分亂碼中字無法正常運作

            print("正在搜尋 : " + All_Stock[_Stock])
            re_do = 0 
            while re_do < 5 :
                try :
                    _stock_Id = All_Stock[_Stock].split('_')[1]

                    Total_Stock_Amounts = Get_Total_Stock_Amounts(_stock_Id)

                    Close_List , Volumn_List = Get_Close_and_volume(_stock_Id , Today_timeformmat , Season_1st_month , Trade_Day_List)

                    tmp_table = pd.read_html("https://fubon-ebrokerdj.fbs.com.tw/z/zc/zcl/zcl.djhtm?a=" + _stock_Id + "&c=" + Season_1st_month.strftime('%Y-%m-%d') + "&d=" + Today_timeformmat.strftime('%Y-%m-%d'))
                    df = tmp_table[2]
                    df = df.drop(df.columns[0:7],axis=0)[:-1]
                    df.columns=["日期","買賣超-外資", "買賣超-投信", "買賣超-自營商", "買賣超-單日合計", "估計持股-外資" ,"估計持股-投信", "估計持股-自營商", "估計持股-單日合計", "持股比重-外資", "持股比重-三大法人"]
                    #特別處理，網站上不知道為啥多出一天非交易日 110/02/09 導致程式錯誤
                    indexDate = df[ df['日期'] == '110/02/09' ].index
                    df.drop(indexDate , inplace=True)
                    #============================================================
                    df = df.reset_index(drop=True)

                    tmp_df_date = []
                    for _tmp in df["日期"] : 
                        tmp_df_date.append(_tmp)

                    Trade_Day_List.reverse() # xxxx/12/31 ~ xxxx/10/01
                    if len(Trade_Day_List) != len(df) : 
                        for check_start in Trade_Day_List :
                            minors_1911 = str(int(check_start.split('/' , 1)[0]) - 1911) + "/" + check_start.split('/' , 1)[1]
                            if minors_1911 not in tmp_df_date : 
                                df.loc[len(df)] = ([minors_1911 , "0" , "0" , "0" , "0" , "0" , "0" , "0" , "0" , "0" , "0" ])
                
                    row_list = [All_Stock[_Stock]]
                    for index , row in df.iterrows() : # index 用不太到，但需要加
                        over_buy = row["買賣超-投信"]
                        if over_buy == "--" : 
                            over_buy = "0"

                        try :
                            Day_percent = str(round(int(over_buy) * 100 / Volumn_List[index] , 3)).replace('-' , "") + "%"
                        except :
                            Day_percent = "0.000%"

                        try :
                            fund_total_percent = str(round(int(row['估計持股-投信']) * 100 / int(Total_Stock_Amounts) , 3))
                            fund_total_percent = fund_total_percent.split("." , 1)[0] + "." + fund_total_percent.split("." , 1)[1].ljust(3 , "0") + "%" # 小數補足三位數
                        except :
                            fund_total_percent = "0.000%"


                        row_list.append(over_buy)
                        row_list.append(Day_percent)
                        row_list.append(fund_total_percent)
                        row_list.append(Close_List[index])
                    
                    if int(row_list[1]) > 0 : # 大於 0 才列入 ,只需要看前面 5 欄 分別是 股票名稱, 買賣超, 當日比, 總持股比, 收盤價
                        Today_overbuy_Stock_list.append([row_list[0] , row_list[1] , row_list[2]])

                    Huge_list.append(row_list)
                    break
                    
                except Exception as e:
                    re_do = re_do + 1

            if re_do == 5 : 
                print(All_Stock[_Stock] + ": 找尋資訊失敗")
                print("https://fubon-ebrokerdj.fbs.com.tw/z/zc/zcl/zcl.djhtm?a=" + _stock_Id + "&c=" + Season_1st_month.strftime('%Y-%m-%d') + "&d=" + Today_timeformmat.strftime('%Y-%m-%d'))
                Error_List.append(All_Stock[_Stock])

    thread_list = []
    for _thread in range (1,51) :
        thread_list.append(threading.Thread( target=do_important_job , args=(_thread , ) ))
    for aa in thread_list:
        aa.start()
    for aa in thread_list:
        aa.join()

    if Error_List != [] : 
        print("以下無法正常寫入檔案 : " + "\n".join(Error_List))
    
    return Huge_list , Today_overbuy_Stock_list

def Count_no_buy_days(main_sheet , sheet_list , Today_overbuy_Stock_list) :

    sheet_data_list = list(map(lambda x: x.title , sheet_list))
    sheet_data_list = list(filter(lambda y : y.find("-Q") != -1 , sheet_data_list))
    
    sheets = [None] * len(sheet_data_list)
    for _s in range (len(sheet_data_list)) :
        sheets[_s] = main_sheet.worksheet_by_title(sheet_data_list[_s])

    result_list = []

    for i in Today_overbuy_Stock_list :
        print(i)

        stop_flag = False
        count_days = 0
        for current_sheet in sheets :

            try :
                row_num = current_sheet.get_col(1 ,include_tailing_empty=False).index(i[0]) + 1 

                if current_sheet.cell('B3').value == "" :
                    specify_row_data_list = current_sheet.get_row(row_num)[5:] # 我會自己手動新增日期
                else :
                    specify_row_data_list = current_sheet.get_row(row_num)[1:] 

                for j in range(0 , len(specify_row_data_list) , 4) :
                    if (int(specify_row_data_list[j]) > 0) :
                        stop_flag = True
                        break
                    else :
                        count_days = count_days + 1

            except : # 當前 sheet 找不到此個股，直接加上 (總寬度-1) /4 的天數
                count_days = count_days + int(len(current_sheet.get_row(3)[5:]) / 4)
            
            if stop_flag == True :
                break
        
        i.append(str(count_days))
        result_list.append(i)
    try : 
        result_list = sorted(result_list , key=lambda xx: int(xx[3]) , reverse=True)
    except :
        None

    return result_list
            




