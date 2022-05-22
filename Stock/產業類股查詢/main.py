from sre_constants import BRANCH
from requests import get
from bs4 import BeautifulSoup
from pandas import read_html
from re import search

res = get("https://www.moneydj.com/z/js/IndustryListNewJS.djjs") 
res.encoding = 'big5'
Industry_list = search ("var NewkindIDNameStr = '(.+)';var" , res.text ).group(1)
Industry_list = Industry_list.split(';')
Industry_dict = {}
for j in Industry_list :
    Main_industry = search(' (.+)~' , j).group(1)
    Sub_industry_list = j.split('~')[1].split(',') # ['C011010 水泥', 'C011011 水泥製品', 'C025012 預拌混凝土', 'C099192 高爐水泥']
    for k in Sub_industry_list :
        Sub_industry_name = k.split(' ')[1] # 水泥
        Sub_industry_code = k.split(' ')[0] # C011010
        Industry_dict[Main_industry + '_' + Sub_industry_name] = Sub_industry_code


All_data = []
for industry , id in Industry_dict.items() :

    try : 
        Stock_List_table = read_html('https://www.moneydj.com/z/zh/zha/ZH00.djhtm?A=' + id)[2] # 第二個表格才是需要的
        Column_Name = Stock_List_table.iloc[1]
        Stock_List_table = Stock_List_table.drop(Stock_List_table.columns[[0,1]],axis=0) # 去除不需要的列
        Stock_List_table.columns = Column_Name # 換 Colume Title
        #Stock_List_table.reset_index(inplace = True , drop=True) 
        Stock_List = Stock_List_table['股票名稱']
        Stock_List = list(map(lambda x : search('\d+(.+)',x).group(1) + '_' + search('\d+',x).group(0)  , Stock_List )) # 修改文字陳述方式 >> 以 XXX_0000 表示
        Stock_List.insert(0, industry) # 在最前面插入產業
        
    except :
        Stock_List = [industry]
    
    print(Stock_List)
    All_data.append(Stock_List)
    