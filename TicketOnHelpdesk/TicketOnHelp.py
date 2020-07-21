# import library to manipulate .msg files
import win32com.client
#Graphic library for interface
import tkinter
# Import library to request the file to manipulate, the askopenfilenames method allows you to manipulate multiple files.
from tkinter.filedialog import askopenfilenames
# Manipulate and fix path system 
from pathlib import PurePath

#https://www.selenium.dev/documentation/es/support_packages/working_with_select_elements/
# Import library for run chrome web driver
from selenium import webdriver
# Import library for select objet in menu or list
from selenium.webdriver.support.select import Select
filedriver = r"C:\Recursos\Programacion\Python\Browser\chromedriver.exe"
# Import library for allow typing on browser tabs.
from selenium.webdriver.common.keys import Keys
# Add a few seconds before writing the ticket subcategory.
import time
# Launch window to select outlook message file(s).
filename = askopenfilenames(initialdir = "/", title = "Select outlook message file")
# Object for manipulate outlook files
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#def GetUserInformation(self, username):
#    pass

driver = webdriver.Chrome(filedriver)

# Browse and manipulate each of the selected message files (Touple Class).
for file in range(len(filename)):
    # Open the outlook element and create a outlook object (using method OpenSharedItem) for manipulate contents. PurePath is used for fix path and Filaename[file] for loop the touple element.
    msg = outlook.OpenSharedItem(PurePath(filename[file]))
    # Initialize the opening of the browser.
    driver.get("http://help")
    driver.get("http://help/NuevoTicket.asp")
    # The title of the ticket is written.
    element = driver.find_element_by_id("H_titulo")
    element.send_keys(msg.Subject)
    # The body of the email is written as the case description to the ticket.
    element = driver.find_element_by_id("H_Descripcion")
    element.send_keys(msg.Body)
    # The category associated with the request is selected.
    select_element = Select(driver.find_element_by_id("H_area"))
    select_element.select_by_visible_text("Solicitudes")
    # Case priority is selected in the ticket (default is low).
    select_element = Select(driver.find_element_by_id("H_prioridad"))
    select_element.select_by_visible_text("Baja")
    # The user's contact mode is selected (default is Email).
    select_element = Select(driver.find_element_by_id("Sel_Contacto"))
    select_element.select_by_visible_text("Email")
    # The user's location is selected.
    select_element = Select(driver.find_element_by_id("H_Ubicacion"))
    select_element.select_by_visible_text("BAQ")
    # The area or department of the user is selected.
    select_element = Select(driver.find_element_by_id("H_OM_dpto"))
    select_element.select_by_visible_text("FINANZAS")
    # The subcategory associated with the request is selected. This field takes time to load the information, 2 seconds of waiting are added before writing the subcategory.
    time.sleep(2) 
    select_element = Select(driver.find_element_by_id("H_SubArea"))
    select_element.select_by_visible_text("Retiro empleado")
    # Open new tab for each processed msg file.
    if (file + 1) < len(filename):
        # Opens a new tab
        driver.execute_script("window.open()")
        # Switch to the newly opened tab
        driver.switch_to.window(driver.window_handles[file+1])