# -*- coding: utf-8 -*-
"""
Created on Fri Apr 24 15:02:32 2020

@author: vevina
"""


from importlib import reload

import win32com.client as win32
import warnings
import pythoncom
import sys
import os
import pandas as pd

reload(sys)

warnings.filterwarnings("ignore")
pythoncom.CoInitialize()
outlook = win32.Dispatch('outlook.application')

def sendmail(receiver, attachment, subject):
    receiver = receiver
    attachment = attachment
    sub = subject
    body = "Please find the attached reports."
    mail = outlook.CreateItem(0)
    mail.To = receiver
    mail.subject = sub.encode('utf-8').decode('utf-8')
    mail.Body = body.encode('utf-8').decode('utf-8')
    for i in range(len(attachment)):
        mail.Attachments.Add(attachment[i])
    mail.Send()


month = input('Month: ')
path = input('Path: ')
addr_path = input('Path of Address: ')
addr = pd.read_excel(addr_path)

for j in range(len(addr.Code)):
    dirlist = []
    sub_code = addr.Code[j].astype("str")
    subjects = sub_code + " - " + month
    for dirpath,dirname,filename in os.walk(path):
        print(addr.Code[j].astype("str"))
        for i in filename:
            if i[0:4] == sub_code:
                dirlist.append(os.path.join(dirpath,i))
                
    sendmail(addr["add"][j], dirlist, subjects)







