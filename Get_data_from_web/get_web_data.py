from selenium import webdriver
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
global web_ids
web_ids = []
def get_web_ids(pth):
    driver.get(pth)
    ids = driver.find_elements(By.TAG_NAME,value='input')
# to get names use '//*[@name]'
    for ii in ids:
        #print (ii)
        if ii.get_attribute('aria-label') is not None:
            print('Tag: ' + ii.get_attribute('aria-label'))
        print('ID: ' + ii.get_attribute('id'))     # element id as string
        if (ii.get_attribute('id') != ''):
            web_ids.append(ii.get_attribute('id'))
        print ('Tag Name: ' + ii.tag_name)
    
    return (web_ids)
pth = 'http://banve.tramthoitiet.vn/#/account/login'
while (len(web_ids) == 0):
    get_web_ids(pth)
print (web_ids)
driver.find_element(by=By.ID,value=web_ids[0]).send_keys('tdbanve')
driver.find_element(by=By.ID,value=web_ids[1]).send_keys('123456aA')
abc = driver.find_elements(By.TAG_NAME,value='button')
for a in abc:
    print ('la: ' + a.get_attribute('type'))
    print ('la: ' + a.get_attribute('id'))
driver.find_elements(By.CSS_SELECTOR,value='button')[0].click()
time.sleep(2)

#driver.find_element(by=By.ID,value='btnSubmit').click()
#driver.maximize_window()
pth1 = 'http://banve.tramthoitiet.vn/#/map/index'
print (pth1)
driver.get(pth1)
get_day = driver.find_elements(By.TAG_NAME,value='input')
for x in get_day:
    print ('input: ' + x.get_attribute('aria-label'))
    print ('input: ' + x.get_attribute('id'))
print ('abc')
time.sleep(5)
#select_element = Select(driver.find_element(by=By.CSS_SELECTOR,value='#DropDownList1'))
#select_element.select_by_value('5')
#time.sleep(10)
#driver.find_element(by=By.ID,value='viewDate').click()
#time.sleep(5)
#driver.find_element(by=By.ID,value='Tab2').click()

#gridview_element = driver.find_element(by="id", value="GridView1")
#rows = gridview_element.find_elements(by=By.TAG_NAME,value="tr")
#data = []
#for row in rows:
#    cells = row.find_elements(by=By.TAG_NAME,value="td")
#    row_data = [cell.text for cell in cells]
#    data.append(row_data)
#df = pd.DataFrame(data)

#print(df)