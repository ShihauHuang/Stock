import pygsheets
from func import datetime , Download_Data_by_date , Get_All_Stock_from_Season_1th_to_Today , Get_fund_buy_info , Count_no_buy_days


def Season( _month ) :
    if _month in ["01" , "02" , "03"] :
        return "Q1"
    elif _month in ["04" , "05" , "06"] :
        return "Q2"
    elif _month in ["07" , "08" , "09"] :
        return "Q3"
    elif _month in ["10" , "11" , "12"] :
        return "Q4" 

# 初始設定 import 這次會使用到的 pygsheets 外，gc 這邊是告訴 Python 我們的授權金鑰 json 放置的位子
gc = pygsheets.authorize(service_file='Fund-37435451c7c1.json') 

# 開啟 GoogleSheet
sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1JJJWutdj7HNOnCR6DrxF6WjIT4IXyOBCd7gK3TMCbLM/')


Today_timeformmat = datetime.datetime.today() # YYYY-MM-DD HH:MM:SS
#Today_timeformmat = datetime.datetime.strptime('20210129' ,'%Y%m%d')
Today_with_Slash = Today_timeformmat.strftime('%Y/%m/%d') # YYYY/MM/DD
Today = Today_timeformmat.strftime('%Y%m%d') # YYYYMMDD


# 下載投信資料
Season_1st_month = Download_Data_by_date(Today_timeformmat) # return 季初的日期 (datetime formmat)

# 將季初到今日投信買過的股票全部放入 list 
All_Stock , Trade_Day_List = Get_All_Stock_from_Season_1th_to_Today(Today_timeformmat , Season_1st_month)
# All_Stock >> ['元大高股息_0056', '中壽_2823', '倉和_6538', '智易_3596'...]
# Trade_Day_List >> ['2020/10/05', '2020/10/06', '2020/10/07', '2020/10/08'...]
Huge_list , Today_overbuy_Stock_list = Get_fund_buy_info(All_Stock , Today_timeformmat , Season_1st_month , Trade_Day_List)

wks_list = sheet.worksheets() # 查看此 GoogleSheet 內 Sheet 清單
n_days_no_buy = Count_no_buy_days(sheet , wks_list , Today_overbuy_Stock_list)


wks = None
for a in wks_list : 
    if a.title == Today[0:4] + "-" + Season(Today[4:6]) : # 已存在
        wks = a
        break
if not wks : # 未存在就新建
    # 首次建立
    wks = sheet.add_worksheet(Today[0:4] + "-" + Season(Today[4:6])) # 新建新表單
    wks.index = 0 # 讓他位置排在最新
    wks.update_value("A1" , "股票名稱") 
    wks.range('A1:A2', returnas='range').merge_cells() # 合併儲存格 A1~A2
    wks.cell("A1").set_vertical_alignment(pygsheets.custom_types.VerticalAlignment.MIDDLE) # 使 A1 上下置中
    wks.delete_cols(2,wks.cols-1) # 先砍掉第一欄以後的欄位，後續會陸續新增
    wks.delete_rows(3,wks.rows-2) # 先砍掉第一列以後的列位，後續會陸續新增 # 要減掉 2 才正常 不知道為啥


# 建立交易日期欄位
Exist_date_list = wks.get_all_values()[0] # 得確認當前已存在的日期 ['股票名稱', '2021/01/05', '', '', '', '2021/01/04'...]
green = (0.8509804, 0.91764706, 0.827451, 0) 
blue = (0.7882353, 0.85490197, 0.972549, 0) 
now_color = green
is_first = False
for trade_day in Trade_Day_List :
    if trade_day not in Exist_date_list : 
        if wks.cols == 1 : # 第一次建立
            wks.add_cols(4)
            is_first = True
        else :
            if wks.cell("B1").color == green : # 再新增欄位前先判斷當前的顏色，再調成另外一個顏色
                now_color = blue
            else :
                now_color = green

            wks.insert_cols(1 , 4) # 在第一欄插入四個空數列，每一天新增
        
        cell_range = [wks.cell('B1')] + wks.range('B2:E2')[0]

        
        for _cell in cell_range :
            _cell.color = now_color # 背景顏色 
            _cell.set_horizontal_alignment(pygsheets.custom_types.HorizontalAlignment.CENTER) # 置中

        # 寫值
        wks.update_values("B1" , [[trade_day , "" , "" , ""] , ["買賣超" , "佔當日成交量 %" , "投信持股比 %" , "收盤價"]] )
        
        # 合併儲存格
        wks.range('B1:E1', returnas='range').merge_cells()
        
        # 調整寬度
        wks.adjust_column_width(start = 2, end = 2, pixel_size = 60) # 買賣超籃為寬度為 60
        wks.adjust_column_width(start = 3, end = 3, pixel_size = 100) # 佔當日成交量 100
        wks.adjust_column_width(start = 4, end = 4, pixel_size = 90) # 投信持股比 60
        wks.adjust_column_width(start = 5, end = 5, pixel_size = 60) # 收盤價 60

if is_first : wks.add_rows(1) # 第一次新建的話要多一行後續直接塞資料


Daily_Sheet = sheet.worksheet_by_title('每日總覽')
empty_list = []
for el in range(Daily_Sheet.rows) :
    empty_list.append(["" , "" , "" , ""])
Daily_Sheet.update_values(crange = "K4"  , values=empty_list) # 先清空
Daily_Sheet.update_values(crange = "K4"  , values=n_days_no_buy) # 在寫入數據


wks.update_values(crange = "A3"  , values=Huge_list)
wks.frozen_rows = 2 # 凍結第二行
wks.frozen_cols = 1 # 凍結第一欄'''
