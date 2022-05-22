import func
import pandas as pd

Stock_Id = "8358"

Total_Stock_Amounts = func.Get_Total_Stock_Amounts(Stock_Id)

tmp_table = pd.read_html("https://fubon-ebrokerdj.fbs.com.tw/z/zc/zcl/zcl.djhtm?a=" + Stock_Id + "&c=2020-10-01&d=2020-12-31")
tmp_table = pd.read_html("https://fubon-ebrokerdj.fbs.com.tw/z/zc/zcl/zcl_8358.djhtm")
df = tmp_table[2]
df = df.drop(df.columns[0:7],axis=0)[:-1]
df.columns=["日期","買賣超-外資", "買賣超-投信", "買賣超-自營商", "買賣超-單日合計", "估計持股-外資" ,"估計持股-投信", "估計持股-自營商", "估計持股-單日合計", "持股比重-外資", "持股比重-三大法人"]
df = df.reset_index(drop=True)

# 投信持股比 % ==============================================================================
for index , row in df.iterrows() : # index 用不太到，但需要加
    # 小數補足兩位數
    ans = str(round(int(row['估計持股-投信']) * 100 / int(Total_Stock_Amounts) , 2))
    ans = ans.split("." , 1)[0] + "." + ans.split("." , 1)[1].ljust(2 , "0")
    print(row["日期"] , ans)
# ==========================================================================================