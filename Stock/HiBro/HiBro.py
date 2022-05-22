from requests import get
from re import findall

# 宏遠 : https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=1260&b=1260&c=E&d=1
# 富邦建國 : https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=9600&b=9658&c=E&d=1
# 統一-內湖 https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=5850&b=0035003800350062&c=E&d=1
# 凱基-信義 https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=9200&b=9216&c=E&d=1
# 凱基-和平 https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=9200&b=9255&c=E&d=1

Dealer_name_Dict = {
    0 : "宏遠" , 
    1 : "富邦建國" ,
    2 : "統一-內湖" ,
    3 : "凱基-信義" ,
    4 : "凱基-和平" ,
}
Dealer_url_Dict = {
    0 : "https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=1260&b=1260&c=E&d=" , # 宏遠
    1 : "https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=9600&b=9658&c=E&d=" , # 富邦建國
    2 : "https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=5850&b=0035003800350062&c=E&d=" , # 統一-內湖
    3 : "https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=9200&b=9216&c=E&d=" , # 凱基-信義
    4 : "https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=9200&b=9255&c=E&d=" , # 凱基-和平
}

while True :
    Stock_list_sort_by_dealer = []

    nearly_num = str(input('最近 N 天 , N = '))
    for i in Dealer_url_Dict : 
        get_data = get(Dealer_url_Dict[i] + nearly_num) # 後面數字為近 N 天
        now_Stock_list = findall("GenLink2stk.+,('.+')" , get_data.text.split('賣超')[0] ) # 1. 會依照買賣超排序下來 2. 過濾賣超，如果要賣超資訊就是[1]
        now_Stock_list = list(map(lambda x: x.replace("'" , '') , now_Stock_list ))
        Stock_list_sort_by_dealer.append(now_Stock_list)
        

    At_least_two_Dealer_buy_Dict = {}
    for major_list in range (len(Stock_list_sort_by_dealer) - 1) :# 從第一項開始比較，最後一個不用比較
        for item in Stock_list_sort_by_dealer[major_list] :

            tmp_list = [] # 建立空白 list
            tmp_list.append(Dealer_name_Dict[major_list]) # 先把自己加進去

            for minor_list in range(major_list + 1 , len(Stock_list_sort_by_dealer)) :
                if item in Stock_list_sort_by_dealer[minor_list] : # 如果此股有能在其他卷商找到
                    tmp_list.append(Dealer_name_Dict[minor_list])
                    del Stock_list_sort_by_dealer[minor_list][Stock_list_sort_by_dealer[minor_list].index(item)] # 刪除相同的，避免重複搜尋
            if len(tmp_list) > 1 :
                At_least_two_Dealer_buy_Dict[item] = tmp_list

    Sort_list = sorted(At_least_two_Dealer_buy_Dict.items() , key = lambda x: len(x[1]) , reverse=True) #由多排到少
    for result in Sort_list :
        print('%s : %s' % (result[0] , ', '.join(result[1]))) 
    print('\n')

