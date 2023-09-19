from selenium import webdriver
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
import pandas as pd
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from datetime import datetime, timedelta
ba_ngay_truoc = str((datetime.now() - timedelta(hours=68)).strftime("%m%d%Y%H"))
date = str((datetime.now() - timedelta(hours=68)).strftime("%m/%d/%Y"))
print (date)
print (ba_ngay_truoc[0:9],str(int(ba_ngay_truoc[9:11])))
options = webdriver.ChromeOptions()
cwd = os.getcwd()
print (cwd)
prefs = {"download.default_directory":cwd}
options.add_experimental_option("prefs",prefs)
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()),options=options)
pth = 'http://222.255.11.82/Modules/MuaTudong/BaoCaoChiTietDuLieu.aspx'
driver.get(pth)
driver.find_element(by=By.ID,value='txtUserName').send_keys('admin')
driver.find_element(by=By.ID,value='txtPWD').send_keys('ttdl@2021')
driver.find_element(by=By.ID,value='btnSubmit').click()
#driver.maximize_window()
pth = 'http://222.255.11.82/Modules/MuaTuDong/frmKhaiThacSoLieuMua.aspx'
driver.get(pth)

select_element = Select(driver.find_element(by=By.CSS_SELECTOR,value='#DropDownList1'))
select_element.select_by_value('5')
time.sleep(10)
bat_dau = driver.find_element(by=By.ID,value='txtOrderDate')
def get_value():
    global value
    value = str(bat_dau.get_attribute("value"))
    return value
date0 = get_value()
while (date0 != date):
    print ('retrying ...')
    bat_dau.send_keys(ba_ngay_truoc[0:9])
    time.sleep(3)
    date0 = get_value()
    print (date0)
select_element = Select(driver.find_element(by=By.ID,value='DropDownListgiodi'))
select_element.select_by_value(str(int(ba_ngay_truoc[9:11])))
driver.find_element(by=By.ID,value='viewDate').click()
time.sleep(5)
driver.find_element(by=By.ID,value='LinkButton1').click()
time.sleep(5)
#os.startfile('ExportExcell.xls')
driver.close()
df = pd.read_html('ExportExcell.xls')
print (df)
df.to_excel('ketqua.xlsx')
os.startfile('ketqua.xlsx')
#print(df)
