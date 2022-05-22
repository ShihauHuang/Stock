import csv
from re import search
import urllib.request
import os.path
import datetime

def Download_Data_by_date (_start_day , _end_day) :
    
    try :
        #print('開始下載資料...')
        twse_url = 'https://www.twse.com.tw/fund/TWT44U?response=csv&date=' #下載網址 上市
        tpex_url = 'https://www.tpex.org.tw/web/stock/3insti/sitc_trading/sitctr_download.php?l=zh-tw&t=D&type=buy&d=' #下載網址 上櫃

        totaldays = (_end_day - _start_day).days + 1

        for daynumber in range(totaldays):
            datestring = (_start_day + datetime.timedelta(days = daynumber)).date()
            Current_day = datestring.strftime("%Y%m%d")
            if not os.path.exists('Data/投信/TWT44U_' + Current_day + '.csv') : # 如果不存在，則下載
                urllib.request.urlretrieve(twse_url + Current_day, "Data/投信/TWT44U_" + Current_day + ".csv")
            if not os.path.exists('Data/投信/sit_' + Current_day + '.csv') : # 如果不存在，則下載
                urllib.request.urlretrieve(tpex_url + str(int(Current_day[0:4]) - 1911) + '/' + Current_day[4:6] + '/' + Current_day[6:8] , "Data/投信/sit_" + Current_day + ".csv")
        #print('下載資料完成.')
    except Exception as e :
        print('下載失敗 : ' + str(e))

def Analysis_data(_start_day , _end_day , _recently_N_days_continuous_buy , enable) :
    Data_dict = {}
    Count = 0 # 用來記錄近幾天的情況，要排除假日
    Current_day = _end_day.strftime("%Y%m%d") 
    
    while Count < _recently_N_days_continuous_buy : # 從後面開始，如果遇到空白的資訊則跳過
        # 上市 ======================================================================================================================================
        with open ('Data/投信/TWT44U_' + Current_day + '.csv' , newline='') as csvfile :
            _rows = csv.reader(csvfile, delimiter=',')
            rows = [_row for _row in _rows]

            if rows[0] != [] : # 上市與上櫃第一行目前來看都是 []，不是 [] 則算入一天
                
                Count = Count + 1
                for row in rows :
                    if row == [] :
                        continue
                    elif row[0] == ' ' :  # 若此判斷成立則為股票欄位
                        Stock_ID = search('(\w+)' , row[1]).group(1)
                        Stock_Name = row[2].strip()
                        Over_Buy = int(int(row[3].replace(',' , '')) / 1000)
                        Over_Sell = int(int(row[4].replace(',' , '')) / 1000)
                        Result = int(int(row[5].replace(',' , '')) / 1000)
                        if Result <= 0 : # 小於 0 以後的數據不用看了
                            break
                        # 第一次較特別
                        if Count == 1 : # 先將第一筆資料全部放入 First 為後續做比對
                            Data_dict[str(Count) + '_' + Stock_ID + '_' + Stock_Name] = [Over_Buy , Over_Sell , Result]
                           
                        else: # 與一開始放進去的資料比對是否已有存在
                            if str(Count - 1) + '_' + Stock_ID + '_' + Stock_Name in Data_dict.keys() : # 如果已經有現有資料，則在 Data 上新增，並相加
                                Data_dict[str(Count) + '_' + Stock_ID + '_' + Stock_Name] = [
                                    Data_dict[str(Count - 1) + '_' + Stock_ID + '_' + Stock_Name][0] + Over_Buy ,
                                    Data_dict[str(Count - 1) + '_' + Stock_ID + '_' + Stock_Name][1] + Over_Sell ,
                                    Data_dict[str(Count - 1) + '_' + Stock_ID + '_' + Stock_Name][2] + Result 
                                ]
        #============================================================================================================================================
        # 上櫃 ======================================================================================================================================   
        with open ('Data/投信/sit_' + Current_day + '.csv' , newline='') as csvfile :
            _rows = csv.reader(csvfile, delimiter=',')
            rows = [_row for _row in _rows]          

            if rows[0] != [] : # 上市與上櫃第一行目前來看都是 []，不是 [] 則算入一天
                #Count = Count + 1 # 基本上上市上櫃情況一樣
                for row in rows :
                    if row == [] :
                        continue
                    elif row[0].isnumeric() :  # 若此判斷成立則為股票欄位
                        Stock_ID = row[1].strip()
                        Stock_Name = row[2].strip()
                        Over_Buy = int(row[3].replace(',' , ''))
                        Over_Sell = int(row[4].replace(',' , ''))
                        Result = int(row[5].replace(',' , ''))
                        # 第一次較特別
                        if Count == 1 : # 先將第一筆資料全部放入 First 為後續做比對
                            Data_dict[str(Count) + '_' + Stock_ID + '_' + Stock_Name] = [Over_Buy , Over_Sell , Result]
                           
                        else: # 與一開始放進去的資料比對是否已有存在
                            if str(Count - 1) + '_' + Stock_ID + '_' + Stock_Name in Data_dict.keys() : # 如果已經有現有資料，則在 Data 上新增，並相加
                                Data_dict[str(Count) + '_' + Stock_ID + '_' + Stock_Name] = [
                                    Data_dict[str(Count - 1) + '_' + Stock_ID + '_' + Stock_Name][0] + Over_Buy ,
                                    Data_dict[str(Count - 1) + '_' + Stock_ID + '_' + Stock_Name][1] + Over_Sell ,
                                    Data_dict[str(Count - 1) + '_' + Stock_ID + '_' + Stock_Name][2] + Result 
                                ]
        #============================================================================================================================================
        Current_day = (datetime.datetime.strptime(Current_day ,'%Y%m%d') - datetime.timedelta(days=1)).date().strftime("%Y%m%d")
    
    if enable == 1 :
        del_list = [] # 先把要刪除的放此，因為 dict 沒法在迴圈時進行刪除
        while datetime.datetime.strptime(Current_day ,'%Y%m%d') >= _start_day : # 檢查前面幾天"都沒有買"的情況

            # 上市 ======================================================================================================================================
            with open ('Data/投信/TWT44U_' + Current_day + '.csv' , newline='') as csvfile :
                
                _rows = csv.reader(csvfile, delimiter=',')
                rows = [_row for _row in _rows]

                for row in rows :
                    if row == [] :
                        continue
                    elif row[0] == ' ' :  # 若此判斷成立則為股票欄位
                        Stock_Name = row[2].strip()
                        Result = int(int(row[5].replace(',' , '')) / 1000)
                        if Result <= 0 : # 小於 0 以後的數據不用看了
                            break
                        for tmp in Data_dict.keys() :
                            if tmp[0].split('_' , 1)[0] == str(Count) and Stock_Name in tmp :
                                del_list.append(tmp)
            #============================================================================================================================================
            # 上櫃 ====================================================================================================================================== 
            with open ('Data/投信/sit_' + Current_day + '.csv' , newline='') as csvfile :
                
                _rows = csv.reader(csvfile, delimiter=',')
                rows = [_row for _row in _rows]

                for row in rows :
                    if row == [] :
                        continue
                    elif row[0].isnumeric() :  # 若此判斷成立則為股票欄位
                        Stock_Name = row[2].strip()
                        Result = int(row[5].replace(',' , ''))
                        if Result <= 0 : # 小於 0 以後的數據不用看了
                            break
                        for tmp in Data_dict.keys() :
                            if tmp[0].split('_' , 1)[0] == str(Count) and Stock_Name in tmp :
                                del_list.append(tmp)        
            #============================================================================================================================================                            
            Current_day = (datetime.datetime.strptime(Current_day ,'%Y%m%d') - datetime.timedelta(days=1)).date().strftime("%Y%m%d")
        del_list = list(set(del_list))
        for _del in del_list :
            del Data_dict[_del]

    print('股票名稱'.ljust(16 , ' ') + '買超數'.ljust(12 , ' ') + '賣超數'.ljust(12 , ' ') + '買賣超'.ljust(12 , ' '))
    Sort_Dict = sorted(Data_dict.items() , key = lambda x: x[1][2] , reverse=True) #由多排到少
    for i in Sort_Dict :
        if i[0].split('_' , 1)[0] == str(Count) :
            Stock_Name = i[0].split('_' , 1)[1]
            space = 0
            for s in Stock_Name.split('_')[1] :
                if len(s.encode('utf-8')) != 1 :
                    space = space + 2
                else :
                    space = space + 1
            Stock_Name_Space = 15 - space
            print(Stock_Name + ' '*Stock_Name_Space + str(i[1][0]).ljust(15 , ' ') + str(i[1][1]).ljust(15 , ' ') + str(i[1][2]).ljust(15 , ' ')  )


while True : 
    ED = str(input('最新日期 (YYYYMMDD) : '))
    SD = str(input('過去日期 (YYYYMMDD) : '))
    recently_N_days_continuous_buy = int(input('連續買超 N 天 , N = '))
    enable = int(input('是否要為近 N 天前開始連買，並且以前沒買過 (0 = 否 ; 1 = 是) : '))

    '''ED = '20201206'
    SD = '20201101'
    recently_N_days_continuous_buy = 2
    enable = 0'''

    start_day_time_format = datetime.datetime.strptime(SD ,'%Y%m%d')
    end_day_time_format = datetime.datetime.strptime(ED ,'%Y%m%d')
    #recently_N_days_continuous_buy = 5 # 此數值為最近三天買，之前沒有買，ex : 20201120~20201130 之間只有 28 29 30 有買超

    Download_Data_by_date(start_day_time_format , end_day_time_format)

    Analysis_data(start_day_time_format , end_day_time_format , recently_N_days_continuous_buy , enable)
    print('====================================================================\n')