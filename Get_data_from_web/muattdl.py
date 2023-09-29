from selenium import webdriver
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
import pandas as pd
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from datetime import datetime, timedelta

start_time = datetime.now()
moment = start_time.strftime("%Y%m%d%H")
print (moment)
if os.path.isfile('token_'+moment+'.txt'):
    print ('Already done, exit')
    exit()
ba_ngay_truoc = str((start_time - timedelta(hours=72)).strftime("%m%d%Y%H"))
long_3ngay_tr = str((start_time - timedelta(hours=72)).strftime("%A, %B %d, %Y"))
date = str((start_time - timedelta(hours=72)).strftime("%m/%d/%Y"))
print (date)
print (ba_ngay_truoc[0:8],str(int(ba_ngay_truoc[8:10])))
options = webdriver.ChromeOptions()
cwd = os.getcwd()
print (cwd)
prefs = {"download.default_directory":cwd}
options.add_experimental_option("prefs",prefs)
options.add_argument('--headless')
#options.add_argument('window-size=360x360')
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()),options=options)
pth = 'http://222.255.11.82/Modules/MuaTudong/BaoCaoChiTietDuLieu.aspx'
driver.get(pth)
driver.find_element(by=By.ID,value='txtUserName').send_keys('admin')
driver.find_element(by=By.ID,value='txtPWD').send_keys('ttdl@2021')
driver.find_element(by=By.ID,value='btnSubmit').click()
#driver.maximize_window()
pth = 'http://222.255.11.82/Modules/MuaTuDong/frmKhaiThacSoLieuMua.aspx'
driver.get(pth)
################
test0 = driver.find_element(by=By.XPATH,value="//*[@id='txtOrderDate']")
test0.click()
for i in range(1,7):
    for j in range(1,8):
        xpath = '/html/body/form/div/table[3]/tbody/tr/td[1]/table[1]/tbody/tr[1]/\
            td[2]/div/div/div[2]/div[1]/table/tbody/tr['+str(i)+']/td['+str(j)+']'
        print (xpath)
        test2 = driver.find_element(by=By.XPATH,value=xpath)
        print (test2.find_element(by=By.TAG_NAME,value='div').get_attribute("title"))
        if (test2.find_element(by=By.TAG_NAME,value='div').get_attribute("title") == long_3ngay_tr):
            test2.click()

#test = driver.find_element(by=By.XPATH,value=xpath)

#for e in test2:
#    print (e.find_element(by=By.TAG_NAME,value='div').get_attribute("title"))
#    if (e.find_element(by=By.TAG_NAME,value='div').get_attribute("title")==long_3ngay_tr):
#        e.click()
################
select_element = Select(driver.find_element(by=By.CSS_SELECTOR,value='#DropDownList1'))
select_element.select_by_value('5')
time.sleep(10)

bat_dau = driver.find_element(by=By.ID,value='txtOrderDate')
def get_bat_dau():
    global value
    value = str(bat_dau.get_attribute("value"))
    return value

date0 = get_bat_dau()
while (date0 != date):
    print ('retrying ...')
    bat_dau.send_keys(ba_ngay_truoc[0:8])
    time.sleep(3)
    date0 = get_bat_dau()
    print (date0)

select_element = Select(driver.find_element(by=By.ID,value='DropDownListgiodi'))
select_element.select_by_value(str(int(ba_ngay_truoc[8:10])))

driver.find_element(by=By.ID,value='viewDate').click()
time.sleep(5)
driver.find_element(by=By.ID,value='LinkButton1').click()
time.sleep(5)
#os.startfile('ExportExcell.xls')
driver.close()
list0 = pd.read_html('ExportExcell.xls',encoding='UTF-8')
os.remove('ExportExcell.xls')
df0 = list0[0]
add_list = list(df0['address'])
#print (add_list)
xa = []
huyen = []
tinh = []
for add in add_list:
    xa.append(add.split(',')[0])
    huyen.append(add.split(',')[1])
    tinh.append(add.split(',')[2])
import numpy as np
xa = np.array(xa)
huyen = np.array(huyen)
tinh = np.array(tinh)
df_sum = pd.DataFrame(df0['StationID'])
df_sum.insert(1,'xa',xa)
df_sum.insert(2,'huyen',huyen)
df_sum.insert(3,'tinh',tinh)

df_sum.insert(4,'Sum6h',df0.iloc[:,-7:-1].sum(axis=1))
df_sum.insert(5,'Sum12h',df0.iloc[:,-13:-1].sum(axis=1))
df_sum.insert(6,'Sum24h',df0.iloc[:,-25:-1].sum(axis=1))
df_sum.insert(7,'Sum48h',df0.iloc[:,-49:-1].sum(axis=1))
df_sum.insert(8,'Sum72h',df0.iloc[:,-73:-1].sum(axis=1))

#print (df_sum)
df_sum = df_sum.sort_values(by=["Sum6h"],ascending=False)
df_sum = df_sum.loc[df_sum["Sum6h"] >= 30]
df_sum["Sum6h"] = df_sum["Sum6h"].apply('{:.1f}'.format)
df_sum["Sum12h"] = df_sum["Sum12h"].apply('{:.1f}'.format)
df_sum["Sum24h"] = df_sum["Sum24h"].apply('{:.1f}'.format)
df_sum["Sum48h"] = df_sum["Sum48h"].apply('{:.1f}'.format)
df_sum["Sum72h"] = df_sum["Sum72h"].apply('{:.1f}'.format)
df_sum.to_excel('Tong_Hop_Mua_BTB.xlsx',engine = 'openpyxl')

df_NA = df_sum.loc[df_sum["tinh"] == " Tỉnh Nghệ An"]
df_TH = df_sum.loc[df_sum["tinh"] == " Tỉnh Thanh Hóa"]
df_HT = df_sum.loc[df_sum["tinh"] == " Tỉnh Hà Tĩnh"]

with pd.ExcelWriter('Tong_Hop_Theo_Tinh.xlsx') as writer:
   df_NA.to_excel(writer,sheet_name='Nghe An')
   df_TH.to_excel(writer,sheet_name='Thanh Hoa')
   df_HT.to_excel(writer,sheet_name='Ha Tinh')

txt6h_NA = '6 giờ qua,\n Tỉnh Nghệ An:\n'
txtsumNA = ''
for i in range(len(df_NA)):
    txt6h_NA = txt6h_NA + df_NA.iloc[i]['xa'] + "," + df_NA.iloc[i]['huyen'] + ": " +\
               str(df_NA.iloc[i]['Sum6h']) +"mm\n"
    txtsumNA = txtsumNA + df_NA.iloc[i]['xa'] + "," + df_NA.iloc[i]['huyen'] + ": Mưa 6h: " +\
               str(df_NA.iloc[i]['Sum6h']) +"mm; Mưa 12h: " + str(df_NA.iloc[i]['Sum12h']) + "mm; Mưa 24h: " +\
               str(df_NA.iloc[i]['Sum24h']) +"mm; Mưa 48h: " + str(df_NA.iloc[i]['Sum48h']) + "mm; Mưa 72h: " +\
               str(df_NA.iloc[i]['Sum72h']) +"mm\n"
if (len(df_NA) >=1):   
    txt6h_NA = '6 giờ qua,\n Tỉnh Nghệ An:\n'
    txtsumNA = '72 giờ qua, \n Tỉnh Nghệ An: \n'
else:
    txt6h_NA = ''
    txtsumNA = ''
if (len(df_NA) >=1):
    for i in range(len(df_NA)):
        txt6h_NA = txt6h_NA + df_NA.iloc[i]['xa'] + "," + df_NA.iloc[i]['huyen'] + ": " +\
               str(df_NA.iloc[i]['Sum6h']) +"mm\n"
        txtsumNA = txtsumNA + df_NA.iloc[i]['xa'] + "," + df_NA.iloc[i]['huyen'] + ": Mưa 6h: " +\
               str(df_NA.iloc[i]['Sum6h']) +"mm; Mưa 12h: " + str(df_NA.iloc[i]['Sum12h']) + "mm; Mưa 24h: " +\
               str(df_NA.iloc[i]['Sum24h']) +"mm; Mưa 48h: " + str(df_NA.iloc[i]['Sum48h']) + "mm; Mưa 72h: " +\
               str(df_NA.iloc[i]['Sum72h']) +"mm\n"
if (len(df_HT) >=1):   
    txt6h_HT = '6 giờ qua,\n Tỉnh Hà Tĩnh:\n'
    txtsumHT = '72 giờ qua, \n Tỉnh Hà Tĩnh: \n'
else:
    txt6h_HT = ''
    txtsumHT = ''
if (len(df_HT) >=1):
    for i in range(len(df_HT)):
        txt6h_HT = txt6h_HT + df_HT.iloc[i]['xa'] + "," + df_HT.iloc[i]['huyen'] + ": " +\
               str(df_HT.iloc[i]['Sum6h']) +"mm\n"
        txtsumHT = txtsumHT + df_HT.iloc[i]['xa'] + "," + df_HT.iloc[i]['huyen'] + ": Mưa 6h: " +\
               str(df_HT.iloc[i]['Sum6h']) +"mm; Mưa 12h: " + str(df_HT.iloc[i]['Sum12h']) + "mm; Mưa 24h: " +\
               str(df_HT.iloc[i]['Sum24h']) +"mm; Mưa 48h: " + str(df_HT.iloc[i]['Sum48h']) + "mm; Mưa 72h: " +\
               str(df_HT.iloc[i]['Sum72h']) +"mm\n"
if (len(df_TH) >=1):
    txt6h_TH = '6 giờ qua,\n Tỉnh Thanh Hóa:\n'
    txtsumTH = '72 giờ qua, \n Tỉnh Thanh Hóa: \n'
else:
    txt6h_TH = ''
    txtsumTH = ''
if (len(df_TH) >=1):
    for i in range(len(df_TH)):
        txt6h_TH = txt6h_TH + df_TH.iloc[i]['xa'] + "," + df_TH.iloc[i]['huyen'] + ":" +\
               str(df_TH.iloc[i]['Sum6h']) +"mm\n"
        txtsumTH = txtsumTH + df_TH.iloc[i]['xa'] + "," + df_TH.iloc[i]['huyen'] + ": Mưa 6h: " +\
               str(df_TH.iloc[i]['Sum6h']) +"mm; Mưa 12h: " + str(df_TH.iloc[i]['Sum12h']) + "mm; Mưa 24h: " +\
               str(df_TH.iloc[i]['Sum24h']) +"mm; Mưa 48h: " + str(df_TH.iloc[i]['Sum48h']) + "mm; Mưa 72h: " +\
               str(df_TH.iloc[i]['Sum72h']) +"mm\n"
### this part send mail
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

def send_mail(send_from,send_to,subject,text,files,username,password):
   msg = MIMEMultipart()
   msg['From'] = send_from
   recipients = send_to
   msg['To'] = ", ".join(recipients)
   msg['Date'] = formatdate(localtime = True)
   msg['Subject'] = subject
   msg.attach(MIMEText(text))
   
   for file in files:
      part = MIMEBase('application', 'octet-stream')
      part.set_payload(open(file, "rb").read())
      encoders.encode_base64(part)
      part.add_header('Content-Disposition', 'attachment', filename=file)
      msg.attach(part)
   
   #context = ssl.SSLContext(ssl.PROTOCOL_SSLv3)
   #SSL connection only working on Python 3+
   smtp = smtplib.SMTP_SSL('smtp.gmail.com', "465", context=context)
   smtp.login(username,password)
   smtp.sendmail(send_from, send_to, msg.as_string())
   smtp.quit()
#########################################

sent_mail = 'bc.nhietmua@gmail.com'
#receive_mail = ['dubaokttv.dbtb@gmail.com','ducbvtvna@gmail.com','tieukttvbtb@gmail.com','tien1967@gmail.com',\
#'huankttv_btb@yahoo.com.vn','leduccuong.kttv@gmail.com','pclbnghean@gmail.com',\
#'radar.vinh.dkvbtb@gmail.com','tangandbkt@gmail.com','nguyenvanluongbtb@gmail.com',\
#'bc.nhietmua@gmail.com','hang3.btb@gmail.com']
name = df0.columns[-2]
print (name)
receive_mail = ['hang3.btb@gmail.com']

subject = 'Báo cáo mưa lớn ngày lúc '+name[2:4]+ ' giờ ngày '+ name[0:2]
body = txt6h_NA + "\n" + txt6h_TH + "\n"+ txt6h_HT+ '\n'+ txtsumNA +"\n"+ txtsumTH+ "\n" + txtsumHT +"\n"
files = ['Tong_Hop_Mua_BTB.xlsx','Tong_Hop_Theo_Tinh.xlsx']
username = 'bc.nhietmua@gmail.com'
password = 'olqx pxbp nvnb eduo'
####

import os
from email.message import EmailMessage
em = EmailMessage()
em['from'] = sent_mail
em['to']= receive_mail
em['subject']=subject
em.set_content(body)

context = ssl.create_default_context()
print ('teste')
send_mail(sent_mail,receive_mail,subject,body,files,username,password)
from shutil import copy2
copy2('token.txt','token_'+moment+'.txt')
print ('pass')