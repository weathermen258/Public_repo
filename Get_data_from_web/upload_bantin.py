from selenium import webdriver
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
import os,sys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from pynput.keyboard import Key, Controller
options = webdriver.ChromeOptions()
cwd = os.getcwd()
print (cwd)
prefs = {"download.default_directory":cwd}
options.add_experimental_option("prefs",prefs)
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()),options=options)
pth = 'http://222.255.11.117:8888/'
driver.get(pth)
khu_vuc = sys.argv[1]
ban_tin = sys.argv[2]
thoi_gian = sys.argv[3]
if khu_vuc == 'NGAN':
    username = 'NGAN'
    password = 'NGAN'
elif khu_vuc == 'BTBO':
    username = 'BTBO'
    password = 'BTBO'
driver.find_element(by=By.ID,value='txt_TaiKhoan').send_keys(username)
driver.find_element(by=By.ID,value='txt_MatKhau').send_keys(password)
driver.find_element(by=By.ID,value='btn_dangky').click()
time.sleep(2)
keyboard = Controller()
keyboard.press(Key.enter)
time.sleep(2)
#driver.maximize_window()
pth = 'http://222.255.11.117:8888/UploadFile.aspx'
driver.get(pth)
time.sleep(5)
select_element = Select(driver.find_element(by=By.ID,value='cphContent_ddl_Kieu'))
if ban_tin == 'DIEM':
    select_element.select_by_value('1')
else:
    select_element.select_by_value('2')
time.sleep(2)
def ma_ban_tin(bantin):
    switcher = {
        'DIEM' : '3',
        'MLDR' : '26',
        'XTND' : '24',
        'MLDL' : '27',
        'KKLR' : '28',
        'NONG' : '29',
        'DONG' : '30',
        'GMTB' : '31',
        }
    return switcher.get(bantin,"Sai bản tin")

select_element = Select(driver.find_element(by=By.ID,value='cphContent_ddl_Loai'))
select_element.select_by_value(ma_ban_tin(ban_tin))
time.sleep(2)

ten_ban_tin = os.path.join(cwd,khu_vuc + '_'+ ban_tin +'_'+ thoi_gian+'.pdf')
ten_ho_so = os.path.join(cwd,'HS_'+khu_vuc + '_'+ ban_tin +'_'+ thoi_gian+'.pdf')
driver.find_element(by=By.ID,value='cphContent_FileUpload1').send_keys(ten_ban_tin)
time.sleep(2)

driver.find_element(by=By.ID,value='cphContent_FileUpload2').send_keys(ten_ho_so)
time.sleep(2)

driver.find_element(by=By.ID,value='cphContent_Button1').click()
time.sleep(2)
keyboard.press(Key.enter)
import ctypes  # An included library with Python install.
def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)
Mbox('Thanh Cong', 'Bản Tin đã upload thành công', 1)




