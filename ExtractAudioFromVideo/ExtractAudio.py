# Tutorial: https://github.com/antonioam82/ejercicios-python/blob/master/ventana_pesta%C3%B1as.py
# Graphic module to request saves and open files.
from tkinter.filedialog import askopenfilename, asksaveasfilename
# Python code to convert video to audio 
import moviepy.editor as mp
import ffmpeg
# Manipulate path system
import os
from pathlib import PurePath

# Launch window to select file.
#filename = PurePath(askopenfilename(initialdir = "/", title = "Select video file", filetypes = (("Video Files", ".mp4"),("All Files","*.*"))))
#filename = PurePath(askopenfilename(initialdir = "/", title = "Select video file"))

# Insert Local Video File Path  
#clip = mp.VideoFileClip(filename)
clip = mp.VideoFileClip(r"C:/test/sample.avi")

# Launch window to save file.
#filename_save = PurePath(asksaveasfilename(initialdir = "/", title = "Save audio file", defaultextension = ".mp3", filetypes = (("Audio Files", ".mp3"),("All Files","*.*"))))

# Insert Local Audio File Path .*mp3
#clip.audio.write_audiofile(filename_save)
clip.audio.write_audiofile(r"C:/test/sample.mp3")
