from selenium import webdriver
import os 
import re
import time
import pandas as pd
from selenium.webdriver.support.ui import Select
import ipywidgets as widgets #套件
from IPython.display import display #顯示
import pandas as pd

def print_value(sender):  
    print(sender.value)
  
number = widgets.Text()
print('欲搜尋幾份ISO標準')
display(number)
number.on_submit(print_value)


ISO_list=[]
In_year_list=[]
for i in range(int(number.value)):
    ISO = widgets.Text()
    print('第'+str(i+1)+'個欲搜尋ISO標準系列')
    display(ISO)
    
    print('第'+str(i+1)+'個欲搜尋該標準幾年內')
    Year = widgets.Text()
    display(Year)
    
    ISO.on_submit(print_value)
    Year.on_submit(print_value)
    ISO_list.append(ISO)
    In_year_list.append(Year)


## 取得當前時間
import time
localtime = time.localtime()
current_year = time.strftime("%Y")
print(type(current_year))


## 爬蟲主程式
if __name__ == '__main__': 
    filepath = os.getcwd() #取得目前工作路徑
    is_success = False   #是否驗證成功的flag
iso_url='https://www.iso.org/committee/45306/x/catalogue/p/1/u/0/w/0/d/0'
get_date=time.strftime("%Y%m%d",time.localtime())
print(get_date)
option = webdriver.ChromeOptions()
prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory':filepath,'profile.default_content_settings.multiple-automatic-downloads': 1}
option.add_experimental_option('prefs',prefs)
option.add_experimental_option('excludeSwitches',['enable-automation'])
driver = webdriver.Chrome(options = option)
driver.get(iso_url)
table = driver.find_element_by_id('datatable-tc-projects')

trlist = driver.find_elements_by_tag_name('tr')
print(len(trlist))
# for i in trlist:
#     tdlist = i.find_elements_by_class_name('sorting_1')
# #     print(tdlist)
#     for col in tdlist:
#         if col.text != '':
#             print(col.text)
# #             print(col.text + '\t',end='')
# #             print('\n')
    


search_iso_result=[]
import re
for i in trlist:
    tdlist = i.find_elements_by_class_name('sorting_1')
#     print(tdlist)
    for col in tdlist:
        for index,year_va in enumerate(In_year_list):
            if col.text != '':
                col_list=col.text.split()
                if len(col_list[1])<4:
                    iso_year_list=re.split(':',col_list[2])
                else :                
                    iso_year_list=re.split(':',col_list[1])
                ##判別搜尋幾年內的標準  
                if  iso_year_list[0].find(ISO_list[index].value) ==0 and int(current_year)-int(iso_year_list[1][:4])<int(year_va.value):
                    print(col.text)
                    search_iso_result.append(col.text)


iso_result_temp=[]
for i,x in enumerate(search_iso_result):
    search_iso_result_dict={}
    temp=re.split("[\r\n]+",x)
    search_iso_result_dict['標準']=temp[0]
    search_iso_result_dict['描述']=temp[1]
    iso_result_temp.append(search_iso_result_dict)


iso_result_df=pd.DataFrame(iso_result_temp)
iso_result_df.to_excel('標準參考.xlsx',index=False)
