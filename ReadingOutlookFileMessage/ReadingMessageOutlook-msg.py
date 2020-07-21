# import library to manipulate .msg files
from outlook_msg import Message
#Graphic library to request the file to manipulate.
import tkinter
from tkinter.filedialog import askopenfilename
# Manipulate path system
import os
from pathlib import PurePath

# Launch window to select file.
filename = PurePath(askopenfilename())

with open(filename) as msg_file:
    msg = Message(msg_file)

# Contents are the plaintext body of the email
contents = msg.body
sender = msg.sender_email
print (contents)
print (sender)


# Use Attachment.save . For example you could do something like this:
#for x in msg_file.attachments:
#    x.save()

# import extract_msg
#msg = extract_msg.Message("path/to/msgfile.msg")
#longFileNames = [x.longFilename for x in msg.attachments] # This is probably the one you want
#shortFileNames = [x.shortFilename for x in msg.attachments]