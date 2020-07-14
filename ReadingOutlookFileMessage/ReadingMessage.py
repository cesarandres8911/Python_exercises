# import library to manipulate .msg files
import win32com.client
#Graphic library to request the file to manipulate.
import tkinter
from tkinter.filedialog import askopenfilename
# Manipulate path system
import os
from pathlib import PurePath

# Launch window to select file.
filename = PurePath(askopenfilename())

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
msg = outlook.OpenSharedItem(filename)

print (msg.SenderName)
print (msg.SenderEmailAddress)
print (msg.SentOn)
print (msg.Subject)
print (msg.Body)

count_attachments = msg.Attachments.Count
if count_attachments > 0:
    for item in range(count_attachments):
        print (msg.Attachments.Item(item + 1).Filename)

#del outlook, msg