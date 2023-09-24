from datetime import datetime, timedelta
ba_ngay_truoc = str((datetime.now() - timedelta(hours=72)).strftime("%m%d%Y%H"))
print (ba_ngay_truoc[0:8],str(int(ba_ngay_truoc[8:10])))

import pandas as pd
list0 = pd.read_html('ExportExcell.xls',encoding='UTF-8')
df0 = list0[0]
name = df0.columns[-2]
print (name)
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
df_sum.to_excel('Tong_Hop_Mua_BTB.xlsx',engine = 'openpyxl')

df_sum = df_sum.loc[df_sum["Sum6h"] >= 30]
df_NA = df_sum.loc[df_sum["tinh"] == " Tỉnh Nghệ An"]
df_TH = df_sum.loc[df_sum["tinh"] == " Tỉnh Thanh Hóa"]
df_HT = df_sum.loc[df_sum["tinh"] == " Tỉnh Hà Tĩnh"]

with pd.ExcelWriter('Tong_Hop_Theo_Tinh.xlsx') as writer:
   df_NA.to_excel(writer,sheet_name='Nghe An')
   df_TH.to_excel(writer,sheet_name='Thanh Hoa')
   df_HT.to_excel(writer,sheet_name='Ha Tinh')
if (len(df_NA) >=1):   
    txt6h_NA = '6 giờ qua,\n Tỉnh Nghệ An:\n'
    txtsumNA = '72 giờ qua, \n Tỉnh Nghệ An: \n'
else:
    txt6h_NA = ''
    txtsumNA = ''
if (len(df_NA) >=1):
    for i in range(len(df_NA)):
        txt6h_NA = txt6h_NA + df_NA.iloc[i]['xa'] + "," + df_NA.iloc[i]['huyen'] + ":" +\
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
        txt6h_HT = txt6h_HT + df_HT.iloc[i]['xa'] + "," + df_HT.iloc[i]['huyen'] + ":" +\
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
receive_mail = ['bc.nhietmua@gmail.com','phannhuxuyen@gmail.com ','phantoantv@gmail.com']
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
print ('pass')
#print (df6h)
#print (df12h)
#print (df72h)
