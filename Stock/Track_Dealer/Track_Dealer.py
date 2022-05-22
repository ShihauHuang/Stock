from requests import get
from re import findall , search
from pandas import read_html
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from configparser import ConfigParser
import webbrowser
from os.path import exists
import threading
from time import sleep
from warnings import filterwarnings
import json

filterwarnings('ignore', category=Warning)

def get_percent_and_rank (All_Stock , Dealer_trade_amount , start_time , end_time) :

    #for Stock in All_Stock :
    def rank_find(Stock) :
        retry = 0
        while retry < 3 :
            try : 
                _stock_id = Stock.split('--')[0]
                _stock_name = Stock.split('--')[1]
                
                for item in Dealer_trade_amount.items() : # 取得資料
                    if item[0].find('Buy') != -1 : 
                        current_name = item[0].split(' Buy ')[1]
                    else :    
                        current_name = item[0].split(' Sell ')[1]
                    if current_name == _stock_name :
                        balance = item[1].split(', ')[2].replace('-' , '')
                        break
                url = 'https://fubon-ebrokerdj.fbs.com.tw/z/zc/zco/zco.djhtm?a=' + _stock_id + '&e=' + start_time + '&f=' + end_time
                while True :
                    res = get(url)
                    if str(res) != '<Response [200]>' :
                        sleep(0.5)
                        continue
                    else : 
                        break
            
                soup = BeautifulSoup(res.content, 'lxml')
                contents = soup.select('table.t01 > tr')
                contents = list(filter(lambda tag : 'id' not in tag.attrs , contents))
                
                own_rank = 0
                for rank in range (len(contents)) : 
                    td_list = contents[rank].select('td')
                    
                    for j in range( 3 , len(td_list) , 5 ) :
                        compare_balance = td_list[j].text.replace(',' , '')
                        if compare_balance == balance : 
                            own_rank = rank + 1
                            break
                    if own_rank != 0 :
                        break
                if own_rank == 0 :
                    own_rank = '大於 15 名'
                else :
                    own_rank = '第 ' + str(own_rank) + ' 名'
                Dealer_trade_amount[item[0]] = item[1] + ', ' + td_list[j + 1].text + ', ' + own_rank
                break
            
            except Exception as e:
                retry = retry + 1
                sleep(0.5)
                continue
        if retry == 3 : 
            Dealer_trade_amount[item[0]] = item[1] + ', --, --'

    thread_list = []
    for _thread in range(len(All_Stock)) :
        thread_list.append(threading.Thread( target=rank_find , args=(All_Stock[_thread] , ) ))
    for aa in thread_list:
        aa.start()
    for aa in thread_list:
        aa.join()

    return Dealer_trade_amount

def date_to_timestamp (_date_string) :
    struct_time = time.strptime(_date_string, "%Y-%m-%d") # 轉成時間元組
    time_stamp = int(time.mktime(struct_time)) # 轉成時間戳
    return str(time_stamp)

def Link_Stock_for_all_dealer_Multiple (event , start_time , end_time) :

    item_ = event.widget.item(event.widget.selection()[0])['values'][0] 
    stock_id = item_.split('--')[0]

    #url = 'https://fubon-ebrokerdj.fbs.com.tw/z/zc/zco/zco_' + stock_id + '.djhtm'
    url = 'https://fubon-ebrokerdj.fbs.com.tw/z/zc/zco/zco.djhtm?a=' + stock_id + '&e=' + start_time + '&f=' + end_time

    chrome_path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"
    if not exists(chrome_path) :
        chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"

    webbrowser.get(chrome_path + ' %s').open(url)

def Link_Stock_by_dealer_Multiple (event , tree ) :

    _dealer_name = event.widget.item(event.widget.selection()[0])['values'][0] 
    dealer_id2 = Hex_format_check(dealer_table[_dealer_name])
    
    item2 = tree.item(tree.selection()[0])['values'][0]
    stock_id = item2.split('--')[0]

    url = 'https://fubon-ebrokerdj.fbs.com.tw/z/zc/zco/zco0/zco0.djhtm?a=' + stock_id + '&b=' + dealer_id2

    chrome_path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"
    if not exists(chrome_path) :
        chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"
        
    webbrowser.get(chrome_path + ' %s').open(url)

def Link_Stock_by_dealer_Single (event , _dealer_name ) :
    
    dealer_id1 = Hex_format_check(dealer_table[_dealer_name])

    _item = event.widget.item(event.widget.selection()[0])['values'][0] # ['2014--中鴻', 36, 0, 36]
    stock_id = _item.split('--')[0]

    url = 'https://fubon-ebrokerdj.fbs.com.tw/z/zc/zco/zco0/zco0.djhtm?a=' + stock_id + '&b=' + dealer_id1

    chrome_path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"
    if not exists(chrome_path) :
        chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"
        
    webbrowser.get(chrome_path + ' %s').open(url)

def Link_Stock_for_all_dealer_Single(event, start_time , end_time) :

    _item = event.widget.item(event.widget.selection()[0])['values'][0] # ['2014--中鴻', 36, 0, 36]
    stock_id = _item.split('--')[0]

    url = 'https://fubon-ebrokerdj.fbs.com.tw/z/zc/zco/zco.djhtm?a=' + stock_id + '&e=' + start_time + '&f=' + end_time

    chrome_path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"
    if not exists(chrome_path) :
        chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"

    webbrowser.get(chrome_path + ' %s').open(url)


def esc_to_destory (event):

    event.widget.winfo_toplevel()
    event.widget.winfo_toplevel().quit() # 讓程式繼續執行
    event.widget.winfo_toplevel().destroy() # 關閉視窗

def Not_in_same_list_need_to_get_data(_stock_id , _dealer_name , _stime , _etime):

    _dealer_id = Hex_format_check(dealer_table[_dealer_name])
    url = 'https://fubon-ebrokerdj.fbs.com.tw/z/zc/zco/zco0/zco0.djhtm?A=' + str(_stock_id) + '&b=' + str(_dealer_id) + '&D=' + _stime + '&E=' + _etime
    _get_data = get(url)
    _get_data.encode = 'utf-8'
    _soup = BeautifulSoup(_get_data.text)
    table2 = _soup.find_all('table')
    df_count = read_html(str(table2))[3]
    df_count = df_count.drop(df_count.columns[[0]],axis=0) # 去除不需要的列
    try : # 有資料情況
        buy_count = str(sum(list(map(lambda x: int(x) , list(df_count[1])))))
        sell_count = str(sum(list(map(lambda x: int(x) , list(df_count[2])))))
        two_minus = str(sum(list(map(lambda x: int(x) , list(df_count[4])))))
        return [_dealer_name , buy_count , sell_count , two_minus]
    except :
        return [_dealer_name , str(0) , str(0) , str(0)]

def browse_detail(event , detail_win , _specify_dealer_list , _Dealer_trade_amount , _stime , _etime) :

    if str(event.widget).find('frame2') != -1 :
        Im = ' Sell '
    else :
        Im = ' Buy '

    try : 
        current_item = event.widget.item(event.widget.selection()[0])['values'] # ['2409--友達', '凱基-信義, 凱基-和平']
        stock_id = current_item[0].split('--')[0] 
        current_stock = current_item[0].split('--')[1] 
        dealers_list = current_item[1].split(', ')

        detail_win.delete(*detail_win.get_children())
        for ddd in _specify_dealer_list : 
            if ddd in dealers_list : 
                tmp_data = _Dealer_trade_amount[ddd + Im + current_stock]
                detail_win.insert("" , tk.END , values = [ddd , tmp_data.split(', ')[0] , tmp_data.split(', ')[1] , tmp_data.split(', ')[2] ])
            else : 
                detail_win.insert("" , tk.END , values = Not_in_same_list_need_to_get_data(stock_id , ddd , _stime , _etime) )
    except :
        print('沒有資料')

def Change_re_string(_str) :
    return _str.replace("*","\*").replace("+","\+").replace("(","\(").replace(")","\)").replace("[","\[").replace("]","\]")

def add_to_list() :
    check_repeat_list = list(tk_List.get(0 , tk.END))
    choose_dealer = dealer_combobox.get().strip()
    if choose_dealer != '' and choose_dealer not in check_repeat_list:
        tk_List.insert(tk.END , dealer_combobox.get())

def Show_Result_Type_Multiple (Buy_Sort_list , Sell_Sort_list , Dealer_trade_amount , _specify_dealer_list) :

    start_time = str(start_date.get()).replace('/' , '-')
    end_time = str(end_date.get()).replace('/' , '-')

    style = ttk.Style()
    style.configure('TNotebook.Tab', font=('URW Gothic L','14','bold'), padding=[10, 10])
    style.configure('Treeview.Heading', font=('URW Gothic L','14','bold'), padding=[10, 10])
    style.configure('Treeview', font=('Calibri','12'), padding=[10, 10])

    result_win = tk.Toplevel()
    result_win.title(', '.join(_specify_dealer_list) + ' 買賣超結果')

    # 買超部分
    Main_Buy_frame = tk.Frame(result_win)

    Main_Buy_Top_frame = tk.Frame(Main_Buy_frame)
    tk.Label(Main_Buy_Top_frame , background = 'Red' , text = "同時買超 (前五十名內)" ,  borderwidth=2, relief="groove" , font = ('Calibri', '15' , 'bold')).pack(anchor=tk.CENTER ,side = tk.TOP, fill = "both")
    Buy_Scrollbar = tk.Scrollbar(Main_Buy_Top_frame)
    Buy_Scrollbar.pack(side = tk.RIGHT , fill = tk.Y)
    columns=("a" , "b")
    Buy_Treeview = ttk.Treeview(Main_Buy_Top_frame , yscrollcommand = Buy_Scrollbar.set , height= 20, show = "headings" , columns=columns) #表格 
    Buy_Treeview.pack(fill = 'both' , pady=(4,0))
    Buy_Scrollbar.config(command = Buy_Treeview.yview)
    Buy_Treeview.column('a',width=180,anchor='w')
    Buy_Treeview.column('b',width=450,anchor='w')
    Buy_Treeview.heading('a', text='股票名稱')
    Buy_Treeview.heading('b', text='卷商名稱')
    for bb in Buy_Sort_list : 
        stock_name = bb[0]
        dealers = ', '.join(bb[1])
        Buy_Treeview.insert("" , tk.END , values = [stock_name,dealers])

    ###
    Main_Buy_Down_frame = tk.Frame(Main_Buy_frame)
    b_d_columns = ('1' , '2' , '3' , '4')
    Buy_Detail_Treeview = ttk.Treeview(Main_Buy_Down_frame , height= len(_specify_dealer_list) , show = "headings" , columns=b_d_columns)
    Buy_Detail_Treeview.pack(fill = 'both' , padx=(4,4) , pady=(0,4))
    Buy_Detail_Treeview.column('1',width=185,anchor='w')
    Buy_Detail_Treeview.column('2',width=155,anchor='e')
    Buy_Detail_Treeview.column('3',width=155,anchor='e')
    Buy_Detail_Treeview.column('4',width=155,anchor='e')
    #Buy_Detail_Treeview.heading('a', text='股票名稱')
    Buy_Detail_Treeview.heading('2', text='買超張數')
    Buy_Detail_Treeview.heading('3', text='賣超張數')
    Buy_Detail_Treeview.heading('4', text='差額')

    # 起始先秀出第一筆 EX : ['1440--南紡', '宏遠, 統一-內湖, 凱基-和平, 富邦-建國']
    try :
        child_id = Buy_Treeview.get_children()[0] # 設定初始點選此項目
        Buy_Treeview.selection_set(child_id) # 設定初始點選此項目
        stock_id = Buy_Treeview.item(child_id)['values'][0].split('--')[0] 
        current_stock = Buy_Treeview.item(child_id)['values'][0].split('--')[1] 
        dealers_list = Buy_Treeview.item(child_id)['values'][1].split(', ')
        for bbd in _specify_dealer_list : 
            if bbd in dealers_list : 
                tmp_data = Dealer_trade_amount[bbd + ' Buy ' + current_stock]
                Buy_Detail_Treeview.insert("" , tk.END , values = [bbd , tmp_data.split(', ')[0] , tmp_data.split(', ')[1] , tmp_data.split(', ')[2] ])
            else : 
                Buy_Detail_Treeview.insert("" , tk.END , values = Not_in_same_list_need_to_get_data(stock_id , bbd , start_time , end_time) )
    except : 
        print('No data')
    ###############################################################################################################################
    # 賣超部分
    Main_Sell_frame = tk.Frame(result_win)

    Main_Sell_Top_frame = tk.Frame(Main_Sell_frame)
    tk.Label(Main_Sell_Top_frame , background = 'Green' , text = "同時賣超 (前五十名內)" ,  borderwidth=2, relief="groove" , font = ('Calibri', '15' , 'bold')).pack(anchor=tk.CENTER ,side = tk.TOP, fill = "both")
    Sell_Scrollbar = tk.Scrollbar(Main_Sell_Top_frame)
    Sell_Scrollbar.pack(side = tk.RIGHT , fill = tk.Y)
    columns=("a1" , "b1")
    Sell_Treeview = ttk.Treeview(Main_Sell_Top_frame , yscrollcommand = Sell_Scrollbar.set , height= 20, show = "headings" , columns=columns) #表格 
    Sell_Treeview.pack(fill = 'both', pady=(4,0))
    Sell_Scrollbar.config(command = Sell_Treeview.yview)
    Sell_Treeview.column('a1',width=180,anchor='w')
    Sell_Treeview.column('b1',width=450,anchor='w')
    Sell_Treeview.heading('a1', text='股票名稱')
    Sell_Treeview.heading('b1', text='卷商名稱')
    for ss in Sell_Sort_list : 
        stock_name = ss[0]
        dealers = ', '.join(ss[1])
        Sell_Treeview.insert("" , tk.END , values = [stock_name,dealers])

    ###
    Main_Sell_Down_frame = tk.Frame(Main_Sell_frame)
    s_d_columns = ('1' , '2' , '3' , '4')
    Sell_Detail_Treeview = ttk.Treeview(Main_Sell_Down_frame , height= len(_specify_dealer_list) , show = "headings" , columns=s_d_columns)
    Sell_Detail_Treeview.pack(fill = 'both' , padx=(4,4) , pady=(0,4))
    Sell_Detail_Treeview.column('1',width=185,anchor='w')
    Sell_Detail_Treeview.column('2',width=155,anchor='e')
    Sell_Detail_Treeview.column('3',width=155,anchor='e')
    Sell_Detail_Treeview.column('4',width=155,anchor='e')
    #Sell_Detail_Treeview.heading('a', text='股票名稱')
    Sell_Detail_Treeview.heading('2', text='買超張數')
    Sell_Detail_Treeview.heading('3', text='賣超張數')
    Sell_Detail_Treeview.heading('4', text='差額')

    # 起始先秀出第一筆 EX : ['1440--南紡', '宏遠, 統一-內湖, 凱基-和平, 富邦-建國']
    try : 
        child_id = Sell_Treeview.get_children()[0] # 設定初始點選此項目
        Sell_Treeview.selection_set(child_id) # 設定初始點選此項目
        stock_id = Sell_Treeview.item(child_id)['values'][0].split('--')[0] 
        current_stock = Sell_Treeview.item(child_id)['values'][0].split('--')[1] 
        dealers_list = Sell_Treeview.item(child_id)['values'][1].split(', ')
        for ssd in _specify_dealer_list : 
            if ssd in dealers_list : 
                tmp_data = Dealer_trade_amount[ssd + ' Sell ' + current_stock]
                Sell_Detail_Treeview.insert("" , tk.END , values = [ssd , tmp_data.split(', ')[0] , tmp_data.split(', ')[1] , tmp_data.split(', ')[2] ])
            else : 
                Sell_Detail_Treeview.insert("" , tk.END , values = Not_in_same_list_need_to_get_data(stock_id , ssd , start_time , end_time) )
    except :
        print('No data')

    ###############################################################################################################################
    Buy_Treeview.bind ('<ButtonRelease-1>' , lambda event , detail_win = Buy_Detail_Treeview , _specify_dealer_list = _specify_dealer_list , _Dealer_trade_amount = Dealer_trade_amount : browse_detail(event , detail_win , _specify_dealer_list , _Dealer_trade_amount , start_time , end_time))
    Buy_Treeview.bind ('<KeyRelease-Down>' , lambda event , detail_win = Buy_Detail_Treeview , _specify_dealer_list = _specify_dealer_list , _Dealer_trade_amount = Dealer_trade_amount : browse_detail(event , detail_win , _specify_dealer_list , _Dealer_trade_amount , start_time , end_time))
    Buy_Treeview.bind ('<KeyRelease-Up>' , lambda event , detail_win = Buy_Detail_Treeview , _specify_dealer_list = _specify_dealer_list , _Dealer_trade_amount = Dealer_trade_amount : browse_detail(event , detail_win , _specify_dealer_list , _Dealer_trade_amount , start_time , end_time))
    Buy_Treeview.bind ('<Double-Button-1>' , lambda event : Link_Stock_for_all_dealer_Multiple(event , start_time , end_time))
    Buy_Treeview.bind ('<Button-3>' , lambda event : Link_Stock_for_all_dealer_Multiple(event , start_time , end_time))
    Sell_Treeview.bind ('<ButtonRelease-1>' , lambda event , detail_win = Sell_Detail_Treeview , _specify_dealer_list = _specify_dealer_list , _Dealer_trade_amount = Dealer_trade_amount : browse_detail(event , detail_win , _specify_dealer_list , _Dealer_trade_amount , start_time , end_time))
    Sell_Treeview.bind ('<KeyRelease-Down>' , lambda event , detail_win = Sell_Detail_Treeview , _specify_dealer_list = _specify_dealer_list , _Dealer_trade_amount = Dealer_trade_amount : browse_detail(event , detail_win , _specify_dealer_list , _Dealer_trade_amount , start_time , end_time))
    Sell_Treeview.bind ('<KeyRelease-Up>' , lambda event , detail_win = Sell_Detail_Treeview , _specify_dealer_list = _specify_dealer_list , _Dealer_trade_amount = Dealer_trade_amount : browse_detail(event , detail_win , _specify_dealer_list , _Dealer_trade_amount , start_time , end_time))
    Sell_Treeview.bind ('<Double-Button-1>' , lambda event : Link_Stock_for_all_dealer_Multiple(event , start_time , end_time))
    Sell_Treeview.bind ('<Button-3>' , lambda event : Link_Stock_for_all_dealer_Multiple(event , start_time , end_time))
    Buy_Detail_Treeview.bind('<Double-Button-1>' , lambda event , tree = Buy_Treeview : Link_Stock_by_dealer_Multiple (event , tree ))
    Buy_Detail_Treeview.bind('<Button-3>' , lambda event , tree = Buy_Treeview : Link_Stock_by_dealer_Multiple (event , tree ))
    Sell_Detail_Treeview.bind('<Double-Button-1>' , lambda event , tree = Sell_Treeview : Link_Stock_by_dealer_Multiple (event , tree ))
    Sell_Detail_Treeview.bind('<Button-3>' , lambda event , tree = Sell_Treeview : Link_Stock_by_dealer_Multiple (event , tree ))

    result_win.bind ('<KeyRelease-Escape>' , esc_to_destory)

    Main_Buy_frame.pack(side = tk.LEFT)
    Main_Buy_Top_frame.pack(side = tk.TOP)
    Main_Buy_Down_frame.pack(side = tk.BOTTOM)

    Main_Sell_frame.pack(side = tk.RIGHT)
    Main_Sell_Top_frame.pack(side = tk.TOP)
    Main_Sell_Down_frame.pack(side = tk.BOTTOM)

    result_win.mainloop()

def Show_Result_Type_Single (Buy_Stock_list , Sell_Stock_list , Dealer_trade_amount , _dealer_name) :

    start_time = str(start_date.get()).replace('/' , '-')
    end_time = str(end_date.get()).replace('/' , '-')

    Dealer_trade_amount = get_percent_and_rank(Buy_Stock_list + Sell_Stock_list , Dealer_trade_amount , start_time , end_time)

    style = ttk.Style()
    style.configure('TNotebook.Tab', font=('URW Gothic L','14','bold'), padding=[10, 10])
    style.configure('Treeview.Heading', font=('URW Gothic L','14','bold'), padding=[10, 10])
    style.configure('Treeview', font=('Calibri','12'), padding=[10, 10])

    result_win = tk.Toplevel()
    result_win.title(_dealer_name + ' 買賣超結果')

    # 買超部分
    Main_Buy_frame = tk.Frame(result_win)
    tk.Label(Main_Buy_frame , background = 'Red' , text = "買超 (前五十名)" ,  borderwidth=2, relief="groove" , font = ('Calibri', '15' , 'bold')).pack(anchor=tk.CENTER ,side = tk.TOP, fill = "both")
    Buy_Scrollbar = tk.Scrollbar(Main_Buy_frame)
    Buy_Scrollbar.pack(side = tk.RIGHT , fill = tk.Y)
    columns=("a" , "b" , "c" , "d" , "e" , "f")
    Buy_Treeview = ttk.Treeview(Main_Buy_frame , yscrollcommand = Buy_Scrollbar.set , height= 28, show = "headings" , columns=columns) #表格 
    Buy_Treeview.pack(fill = 'both' , padx=(4,4) , pady=(4,4))
    Buy_Scrollbar.config(command = Buy_Treeview.yview)
    Buy_Treeview.column('a',width=185,anchor='w')
    Buy_Treeview.column('b',width=115,anchor='e')
    Buy_Treeview.column('c',width=115,anchor='e')
    Buy_Treeview.column('d',width=115,anchor='e')
    Buy_Treeview.column('e',width=135,anchor='e')
    Buy_Treeview.column('f',width=115,anchor='e')
    Buy_Treeview.heading('a', text='股票名稱')
    Buy_Treeview.heading('b', text='買超張數')
    Buy_Treeview.heading('c', text='賣超張數')
    Buy_Treeview.heading('d', text='差額')
    Buy_Treeview.heading('e', text='當日成交比')
    Buy_Treeview.heading('f', text='買超名次')
    for bb in Buy_Stock_list :  #['2888--新光金', '2303--聯電', '00677U--富邦VIX', '3481--群創', '2448--晶電', '1438--三地開發', '2409--友達', '2449--京元電子', '1806--冠軍']
        stock_name = bb.split('--')[1]
        tmp_data = Dealer_trade_amount[_dealer_name + ' Buy ' + stock_name]
        Buy_Treeview.insert("" , tk.END , values = [bb , tmp_data.split(', ')[0] , tmp_data.split(', ')[1] , tmp_data.split(', ')[2] ,  tmp_data.split(', ')[3] ,  tmp_data.split(', ')[4]])

    # 賣超部分
    Main_Sell_frame = tk.Frame(result_win)
    tk.Label(Main_Sell_frame , background = 'Green' , text = "賣超 (前五十名)" ,  borderwidth=2, relief="groove" , font = ('Calibri', '15' , 'bold')).pack(anchor=tk.CENTER ,side = tk.TOP, fill = "both")
    Sell_Scrollbar = tk.Scrollbar(Main_Sell_frame)
    Sell_Scrollbar.pack(side = tk.RIGHT , fill = tk.Y)
    columns=("a" , "b" , "c" , "d" , "e" , "f")
    Sell_Treeview = ttk.Treeview(Main_Sell_frame , yscrollcommand = Sell_Scrollbar.set , height= 28, show = "headings" , columns=columns) #表格 
    Sell_Treeview.pack(fill = 'both' , padx=(4,4) , pady=(4,4))
    Sell_Scrollbar.config(command = Sell_Treeview.yview)
    Sell_Treeview.column('a',width=185,anchor='w')
    Sell_Treeview.column('b',width=115,anchor='e')
    Sell_Treeview.column('c',width=115,anchor='e')
    Sell_Treeview.column('d',width=115,anchor='e')
    Sell_Treeview.column('e',width=135,anchor='e')
    Sell_Treeview.column('f',width=115,anchor='e')
    Sell_Treeview.heading('a', text='股票名稱')
    Sell_Treeview.heading('b', text='買超張數')
    Sell_Treeview.heading('c', text='賣超張數')
    Sell_Treeview.heading('d', text='差額')
    Sell_Treeview.heading('e', text='當日成交比')
    Sell_Treeview.heading('f', text='賣超名次')
    for ss in Sell_Stock_list :  #['2888--新光金', '2303--聯電', '00677U--富邦VIX', '3481--群創', '2448--晶電', '1438--三地開發', '2409--友達', '2449--京元電子', '1806--冠軍']
        stock_name = ss.split('--')[1]
        tmp_data = Dealer_trade_amount[_dealer_name + ' Sell ' + stock_name]
        Sell_Treeview.insert("" , tk.END , values = [ss , tmp_data.split(', ')[0] , tmp_data.split(', ')[1] , tmp_data.split(', ')[2] ,  tmp_data.split(', ')[3] ,  tmp_data.split(', ')[4]])
    
    #####################################################
    result_win.bind ('<KeyRelease-Escape>' , esc_to_destory)
    Buy_Treeview.bind('<Double-Button-1>' , lambda event , _dealer_name = _dealer_name : Link_Stock_by_dealer_Single (event , _dealer_name ) )
    Buy_Treeview.bind('<Return>' , lambda event , _dealer_name = _dealer_name : Link_Stock_by_dealer_Single (event , _dealer_name ) )
    Buy_Treeview.bind('<Button-3>' , lambda event : Link_Stock_for_all_dealer_Single (event , start_time , end_time ) )
    Sell_Treeview.bind('<Double-Button-1>' , lambda event , _dealer_name = _dealer_name : Link_Stock_by_dealer_Single (event , _dealer_name ) )
    Sell_Treeview.bind('<Return>' , lambda event , _dealer_name = _dealer_name : Link_Stock_by_dealer_Single (event , _dealer_name ) )
    Sell_Treeview.bind('<Button-3>' , lambda event : Link_Stock_for_all_dealer_Single (event , start_time , end_time ) )

    Main_Buy_frame.pack(side = tk.LEFT)
    Main_Sell_frame.pack(side = tk.RIGHT)

    result_win.mainloop()

def Run_Main(_specify_dealer_list , _paramter_list) :

    Buy_Stock_list_sort_by_dealer = []
    Sell_Stock_list_sort_by_dealer = []
    Dealer_trade_amount = {}

    # <a href="javascript:Link2Stk('00637L');">00637L元大滬深300正2</a>
    # GenLink2stk('AS9103','美德醫療-DR');
    search_pattern1 = "(?:Link2Stk.+?>|GenLink2stk.+?,')(.+)(?:<|'\))"
    search_pattern2 = "(?:Link2Stk\('|GenLink2stk\('\D+)(.+)(?:'\);\"|',)"    
    Overbuy_pattern ='(?:\'|<)(?:\s+?|.+?)+?nowrap>([\d,]+)' 
    Oversell_pattern = '(?:\'|<)(?:\s+?|.+?)+?nowrap(?:\s+?|.+?)+?nowrap>([\d,]+)'     

    for _p in _paramter_list :
        
        current_dealer = _specify_dealer_list[_paramter_list.index(_p)]

        get_data = get("https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?" + _p)
        get_data.encode = 'utf-8'
    
        Buy_Stock_Name_list = findall(search_pattern1 , get_data.text.split('賣超')[0] ) # 1. 會依照買賣超排序下來 2. 過濾賣超，如果要賣超資訊就是[1]
        Buy_Stock_ID_list = findall(search_pattern2 , get_data.text.split('賣超')[0] )
        Buy_Stock_list = list(map(lambda x , y: y + '--' + x.replace(y , '') , Buy_Stock_Name_list , Buy_Stock_ID_list))
        Buy_Stock_list_sort_by_dealer.append(Buy_Stock_list)
        for _bs in Buy_Stock_list :
            bs = _bs.split('--')[1]
            Buy_OverBuy = search(Change_re_string(bs) + Overbuy_pattern , get_data.text.split('賣超')[0] ).group(1).replace(',' , '')
            Buy_OverSell = search(Change_re_string(bs) + Oversell_pattern , get_data.text.split('賣超')[0] ).group(1).replace(',' , '')
            Buy_result = str(int(Buy_OverBuy) - int(Buy_OverSell))
            Dealer_trade_amount[ current_dealer + ' Buy ' + bs] = Buy_OverBuy + ', ' + Buy_OverSell + ', ' + Buy_result

        Sell_Stock_Name_list = findall(search_pattern1 , get_data.text.split('賣超')[1] )
        Sell_Stock_ID_list = findall(search_pattern2 , get_data.text.split('賣超')[1] )
        Sell_Stock_list = list(map(lambda x , y: y + '--' + x.replace(y , '') , Sell_Stock_Name_list , Sell_Stock_ID_list))
        Sell_Stock_list_sort_by_dealer.append(Sell_Stock_list)
        for _bs1 in Sell_Stock_list :
            bs1 = _bs1.split('--')[1]
            Sell_OverBuy = search(Change_re_string(bs1) + Overbuy_pattern , get_data.text.split('賣超')[1] ).group(1).replace(',' , '')
            Sell_OverSell = search(Change_re_string(bs1) + Oversell_pattern , get_data.text.split('賣超')[1] ).group(1).replace(',' , '')
            Sell_result = str(int(Sell_OverBuy) - int(Sell_OverSell))
            Dealer_trade_amount[ current_dealer + ' Sell ' + bs1] = Sell_OverBuy + ', ' + Sell_OverSell + ', ' + Sell_result

    if len(_specify_dealer_list) >= 2 : # 使用者選 2 個以上的卷商情況，至少同時要出現兩個，若只有選一個則列出自己
        At_least_two_Dealer_buy_Dict = {}
        for b_major_list in range (len(Buy_Stock_list_sort_by_dealer) - 1) :# 從第一項開始比較，最後一個不用比較
            for item in Buy_Stock_list_sort_by_dealer[b_major_list] :
                b_tmp_list = [] # 建立空白 list
                b_tmp_list.append(_specify_dealer_list[b_major_list]) # 先把自己加進去

                for b_minor_list in range(b_major_list + 1 , len(Buy_Stock_list_sort_by_dealer)) :
                    if item in Buy_Stock_list_sort_by_dealer[b_minor_list] : # 如果此股有能在其他卷商找到
                        b_tmp_list.append(_specify_dealer_list[b_minor_list])
                        del Buy_Stock_list_sort_by_dealer[b_minor_list][Buy_Stock_list_sort_by_dealer[b_minor_list].index(item)] # 刪除相同的，避免重複搜尋
                if len(b_tmp_list) > 1 :
                    At_least_two_Dealer_buy_Dict[item] = b_tmp_list
        Buy_Sort_list = sorted(At_least_two_Dealer_buy_Dict.items() , key = lambda x: len(x[1]) , reverse=True) #由多排到少


        At_least_two_Dealer_sell_Dict = {}
        for s_major_list in range (len(Sell_Stock_list_sort_by_dealer) - 1) :# 從第一項開始比較，最後一個不用比較
            for item1 in Sell_Stock_list_sort_by_dealer[s_major_list] :
                s_tmp_list = [] # 建立空白 list
                s_tmp_list.append(_specify_dealer_list[s_major_list]) # 先把自己加進去

                for s_minor_list in range(s_major_list + 1 , len(Sell_Stock_list_sort_by_dealer)) :
                    if item1 in Sell_Stock_list_sort_by_dealer[s_minor_list] : # 如果此股有能在其他卷商找到
                        s_tmp_list.append(_specify_dealer_list[s_minor_list])
                        del Sell_Stock_list_sort_by_dealer[s_minor_list][Sell_Stock_list_sort_by_dealer[s_minor_list].index(item1)] # 刪除相同的，避免重複搜尋
                if len(s_tmp_list) > 1 :
                    At_least_two_Dealer_sell_Dict[item1] = s_tmp_list
        Sell_Sort_list = sorted(At_least_two_Dealer_sell_Dict.items() , key = lambda x: len(x[1]) , reverse=True) #由多排到少

        #####
        Show_Result_Type_Multiple (Buy_Sort_list , Sell_Sort_list , Dealer_trade_amount , _specify_dealer_list)
        #####

    else : # 做單一
        #####
        Show_Result_Type_Single (Buy_Stock_list_sort_by_dealer[0] , Sell_Stock_list_sort_by_dealer[0] , Dealer_trade_amount , _specify_dealer_list[0])
        #####

def Hex_format_check(_id) :
    
    result = ''

    try : 
        int(_id)
        result = _id
    
    except : # 若不為純數字需要轉 16 進制
        for x in _id :
            tmp_string = ord(x)
            tmp_string = '00' + hex(tmp_string)[2:]
            result += tmp_string
    return result

def search_data() :

    dealer_id_list = [] # 結果放入於此
    specify_dealer_list = list(tk_List.get(0 , tk.END))

    if len(specify_dealer_list) != 0 : # 有東西才做

        for _name in specify_dealer_list :

            #paramter_b = dealer_table[dealer_table[1].isin([_name])].iat[0,0] # 查找語法
            paramter_b = dealer_table[_name]

            if _name.find('-') != -1 : # 有找到 - 代表為分行
                main_company_name = _name.split('-')[0].strip()
                if main_company_name == "臺銀" : # 例外
                    main_company_name = "臺銀證券"
                elif main_company_name == "臺灣企銀" : # 例外
                    main_company_name = "台灣企銀"
                paramter_a = dealer_table[main_company_name]
            else :
                paramter_a = paramter_b

            paramter_a = Hex_format_check(paramter_a)
            paramter_b = Hex_format_check(paramter_b)

            # a=9200&b=9236&c=E&e=2020-11-20&f=2020-11-20
            final_paramter = 'a=' + paramter_a + '&b=' + paramter_b + '&c=E&e=' + str(start_date.get()).replace('/' , '-') + '&f=' + str(end_date.get()).replace('/' , '-')
            dealer_id_list.append(final_paramter)
        
        #########
        Run_Main(specify_dealer_list , dealer_id_list)
        #########



def Save_list_event(event) :
    conf = ConfigParser()
    conf.read('Config.cfg',encoding='utf-8')
    Key = event.keysym.lower()
    if Key in conf.options('Define') : # 限制只使用 F1~F12
        save_list = list(tk_List.get(0 , tk.END))
        save_info = ', '.join(save_list)
        conf.set('Define' , Key , save_info)
        conf.write(open('Config.cfg','w',encoding='utf-8'))

def Load_list_or_delete_event(event) :

    if event.keysym == "BackSpace" or event.keysym == "Delete" :
        del_num_list = list(tk_List.curselection())
        del_num_list.reverse()
        for d in del_num_list :
            tk_List.delete(d)
    else :
        conf = ConfigParser()
        conf.read('Config.cfg',encoding='utf-8')
        Key = event.keysym.lower()
        if Key in conf.options('Define') : # 限制只使用 F1~F12
            tk_List.delete(0,tk.END) # 先清除再 load
            load_info = conf.get('Define' , Key)
            if load_info != '' :
                load_list = load_info.split(', ')            
                for l in load_list :
                    tk_List.insert(tk.END , l)

#https://www.twse.com.tw/zh/brokerService/brokerServiceAudit

dealer_table = {}

res = get('https://openapi.twse.com.tw/v1/opendata/t187ap18')  #證券商基本資料
res = json.loads(res.text)
for dd in res :
    if dd['證券代號'][0:3] == '601' : #('亞證券', '6010') ('\uedac亞-網路', '6012') ('\uee7d?亞-鑫豐', '601d')
        txt = search( '(亞.+)' ,  dd['券商(證券IB)簡稱'] ).group(1)
        dd['券商(證券IB)簡稱'] = '(牛牛牛)' + txt
    if dd['券商(證券IB)簡稱'].replace(' ','').find('自營') == -1 :
        dealer_table[dd['券商(證券IB)簡稱'].replace(' ','')] = dd['證券代號']

res = get('https://openapi.twse.com.tw/v1/opendata/OpenData_BRK02')  #證券商分公司基本資料
res = json.loads(res.text)
for dd in res :
    if dd['證券商代號'][0:3] == '601' : #('亞證券', '6010') ('\uedac亞-網路', '6012') ('\uee7d?亞-鑫豐', '601d')
        txt = search( '(亞.+)' ,  dd['證券商名稱'] ).group(1)
        dd['證券商名稱'] = '(牛牛牛)' + txt
    if dd['證券商名稱'].replace(' ','').find('自營') == -1 :
        dealer_table[dd['證券商名稱'].replace(' ','')] = dd['證券商代號']

dealer_table = sorted( dealer_table.items() , key=lambda x:x[1] ) # 變成 list 型態
dealer_table = dict( (x, y) for x, y in dealer_table ) # 轉回去 dict {'合庫': '1020', '合庫-台中': '1021'...}


app = tk.Tk() 
app.title(u'卷商同買賣追蹤')
app.resizable (False , False) # 不可縮放

Combobox_text_font = ('Calibri', '24')
Combobox_list_text_font = ('Courier New', '20')

app.option_add('*TCombobox*Listbox.font', Combobox_list_text_font)   # apply font to combobox list

Label_Top = tk.Label(app , font = ('Calibri', '20') , text = "選擇卷商")
Label_Top.grid( row = 0 , column = 0 , columnspan = 19 , sticky=tk.W+tk.E+tk.N+tk.S , padx='5', pady='3')

dealer_combobox = ttk.Combobox(app , state='readonly' , font = Combobox_text_font  , values = list(dealer_table.keys()) )
dealer_combobox.grid( row = 1 , column = 0 , columnspan = 17, sticky=tk.W+tk.E+tk.N+tk.S  , padx=(5,0.5), pady='3')

add_btn = tk.Button(app , font = ('Calibri', '20') , text = '加入' , command = add_to_list)
add_btn.grid( row = 1 , column = 17 , sticky=tk.W+tk.E+tk.N+tk.S  , padx=(0,5.5), pady='3')

tk_List = tk.Listbox(app , selectmode= tk.EXTENDED , font = ('Calibri', '16'))
tk_List.grid( row = 2 , column = 0 , columnspan = 19 , sticky=tk.W+tk.E+tk.N+tk.S  , padx='5', pady='3')
tk_List.bind('<Key>' , Load_list_or_delete_event)
tk_List.bind("<Control-Key>", Save_list_event)

start_date = DateEntry(app, date_pattern='yyyy/mm/dd')
start_date.grid( row = 3 , column = 0 , columnspan = 10 , sticky=tk.W+tk.E+tk.N+tk.S  , padx=(5,0), pady='3')

Label_Middle = tk.Label(app , font = ('Calibri', '20') , text = "～")
Label_Middle.grid( row = 3 , column = 10 , sticky=tk.W+tk.E+tk.N+tk.S  , pady='3')

end_date = DateEntry(app, date_pattern='yyyy/mm/dd')
end_date.grid( row = 3 , column = 11 , columnspan = 8 , sticky=tk.W+tk.E+tk.N+tk.S  , padx=(0,5), pady='3')

search_btn = tk.Button(app , font = ('Calibri', '20') , text = '查詢' , command = search_data)
search_btn.grid( row = 4 , column = 0 , columnspan = 19 , sticky=tk.W+tk.E+tk.N+tk.S  , padx='5', pady='3')


app.update()
#x_axis = int((app.winfo_screenwidth() - app.winfo_width()) / 2) # 修正起始出現位置
y_axis = int((app.winfo_screenheight() - app.winfo_height()) / 2)
app.geometry('430x498+200+' + str(y_axis - int(app.winfo_width() / 3) ))

app.mainloop()
