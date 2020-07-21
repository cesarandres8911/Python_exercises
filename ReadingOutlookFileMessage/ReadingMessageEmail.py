# import library to manipulate .msg files
import email
#Graphic library to request the file to manipulate.
import tkinter
from tkinter.filedialog import askopenfilename
# Manipulate path system
import os
from pathlib import PurePath

# Launch window to select file.
filename = PurePath(askopenfilename())

#b = email.message_from_binary_file(filename)

