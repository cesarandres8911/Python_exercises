# https://www.selenium.dev/documentation/es/webdriver/browser_manipulation/
# https://selenium-python.readthedocs.io/installation.html
import win32com.client                                                          # import library to manipulate .msg files
import tkinter                                                                  #Graphic library for interface
from tkinter.filedialog import askopenfilenames                                 # Import library to request the file to manipulate, the askopenfilenames method allows you to manipulate multiple files.
from pathlib import PurePath                                                    # Manipulate and fix path system
from selenium import webdriver                                                  # Import library for run chrome web driver (https://www.selenium.dev/documentation/es/support_packages/working_with_select_elements/)
from selenium.webdriver.support.select import Select                            # Import library for select objet in menu or list
from selenium.webdriver.common.keys import Keys                                 # Import library for allow typing on browser tabs.
from selenium.webdriver.chrome.options import Options                           # Libs for edit option in chrome drive.
from selenium.webdriver import ActionChains                                     # Method for manipulate clic or action on button.
import time                                                                     # Add a few seconds before writing the ticket subcategory.

filename = askopenfilenames(title = "Select outlook message file")              # Launch window to select outlook message file(s).
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")  # Object for manipulate outlook files
filedriver = r"C:\Recursos\Programacion\Python\Browser\chromedriver.exe"        # Path for chromedriver file to run browser.
options = Options()                                                             # Create option object for modify argument to chrome driver 
options.add_argument("--start-maximized")                                       # code for maximized windows to start chrome driver
driver=webdriver.Chrome(options=options, executable_path=filedriver)     # Initialize the chrome browser and maximize windows.

for file in range(len(filename)):                                               # Browse and manipulate each of the selected message files (Touple Class).
    msg = outlook.OpenSharedItem(PurePath(filename[file]))                      # Open the outlook element and create a outlook object (using method OpenSharedItem) for manipulate contents. PurePath is used for fix path and Filaename[file] for loop the touple element.
    driver.get("http://help")                                                   # Initialize the opening of the browser.
    driver.get("http://help/NuevoTicket.asp")
    element = driver.find_element_by_name("usu_nombre")
    element.send_keys("cplata")
    element.send_keys(Keys.TAB)
    element = driver.find_element_by_id("H_titulo")                             # The title of the ticket is written.
    element.send_keys(msg.Subject)
    element = driver.find_element_by_id("H_Descripcion")                        # The body of the email is written as the case description to the ticket.
    bodysplit = msg.Body.split()                                                # Delete line breaks from message body
    element.send_keys(' '.join(bodysplit))                                      # Formatting the body of the message to avoid typing errors in the help application.
    select_element = Select(driver.find_element_by_id("H_area"))                # The category associated with the request is selected.
    select_element.select_by_visible_text("Solicitudes")
    select_element = Select(driver.find_element_by_id("H_prioridad"))           # Case priority is selected in the ticket (default is low).
    select_element.select_by_visible_text("Baja")
    select_element = Select(driver.find_element_by_id("Sel_Contacto"))          # The user's contact mode is selected (default is Email).
    select_element.select_by_visible_text("Email")
    select_element = Select(driver.find_element_by_id("H_Ubicacion"))           # The user's location is selected.
    select_element.select_by_visible_text("CAL")
    select_element = Select(driver.find_element_by_id("H_OM_dpto"))             # The area or department of the user is selected.
    select_element.select_by_visible_text("GESTION HUMANA") 
    button = driver.find_element_by_xpath("//*[@id='adjuntar_archivo']")
    ActionChains(driver).click(button).perform()
    driver.switch_to.window(driver.window_handles[file+1])
    element = driver.find_element_by_xpath("//*[@id='file']")
    element.send_keys(filename[file])
    button = driver.find_element_by_xpath("//*[@id='enviar']")
    ActionChains(driver).click(button).perform()

    driver.switch_to.alert.accept()
    driver.switch_to.window(driver.window_handles[file])

    time.sleep(2)                                                               # This field takes time to load the information, 2 seconds of waiting are added before writing the subcategory.
    select_element = Select(driver.find_element_by_id("H_SubArea"))             # The subcategory associated with the request is selected.
    select_element.select_by_visible_text("Retiro empleado")
    
    if (file + 1) < len(filename):                                              # Open new tab for each processed msg file.
        driver.execute_script("window.open()")                                  # Opens a new tab
        driver.switch_to.window(driver.window_handles[file+1])                  # Switch to the newly opened tab


def OpenBrowser(parameter_list):
    pass

def CreateTicket(parameter_list):
    pass