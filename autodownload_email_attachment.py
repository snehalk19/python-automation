'''Everyday there are many mails that are piled up in our mailbox.
This script reduces the time of opening the mail and downloading the attachments everyday.
We just need to provide the subject of mail,the attachment file name and the location where the attached file needs to be saved
This script is written for windows OS and outlook Application '''

import inbox as inbox
import win32com.client
from win32com.client import Dispatch
import datetime as date
import os.path

def attach(subject, name):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox_1 = outlook.Folders[0]
    print("Inbox",inbox_1)

    subfolder = inbox_1.Folders[1]
    print("Subfolder",subfolder)

    inbox = subfolder.Folders[1]
    print(inbox)

    all_inbox = inbox.Items
    val_date = date.date.today()
    sub_today = subject
    att_today = name
    for msg in all_inbox:
        print("msg",msg.Subject)
        if msg.Subject == sub_today:
            break
    for att in msg.Attachments:
        print("attr",att)
        if att.FileName == att_today:
            break
    att.SaveASFile('D:\save' + '\\' + att.FileName)
    print("Mail Successfully Extracted")


attach('A',"B")