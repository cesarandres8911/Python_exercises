# https://www.selenium.dev/documentation/es/webdriver/browser_manipulation/
# https://selenium-python.readthedocs.io/installation.html
# https://docs.hektorprofe.net/python/funcionalidades-avanzadas/expresiones-regulares/
from tkinter.constants import FALSE
import win32com.client                                                          # import library to manipulate .msg files
import tkinter                                                                  # Graphic library for interface
from tkinter.filedialog import askopenfilenames                                 # Import library to request the file to manipulate, the askopenfilenames method allows you to manipulate multiple files.
from pathlib import PurePath                                                    # Manipulate and fix path system
from selenium import webdriver                                                  # Import library for run chrome web driver (https://www.selenium.dev/documentation/es/support_packages/working_with_select_elements/)
from selenium.webdriver.support.select import Select                            # Import library for select objet in menu or list
from selenium.webdriver.common.keys import Keys                                 # Import library for allow typing on browser tabs.
from selenium.webdriver.chrome.options import Options                           # Libs for edit option in chrome drive.
from selenium.webdriver import ActionChains                                     # Method for manipulate clic or action on button.
import time                                                                     # Add a few seconds before writing the ticket subcategory.
from configparser import ConfigParser                                           # Manipulate the local system variables
import re

ConfigFile = ConfigParser()
ConfigFile.read("config.ini")
UserInfo = ConfigFile["USERINFO"]
SiteInfo = ConfigFile["HELP"]
filename = askopenfilenames(title = "Select outlook message file")              # Launch window to select outlook message file(s).
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")  # Object for manipulate outlook files
filedriver = UserInfo["chromedriverpath"]                                       # Path for chromedriver file to run browser.
options = Options()                                                             # Create option object for modify argument to chrome driver 
options.add_argument("--start-maximized")                                       # Maximized windows to start chrome driver
driver=webdriver.Chrome(options=options, executable_path=filedriver)            # Initialize the chrome browser 

def IsRetiredUser(pos):
    """ Check if it is a user retired email

    Search the message title for the words associated with a user retired email.

    Parameters

    pos -- email message position in the tuple class
    """
    msg = outlook.OpenSharedItem(PurePath(filename[pos])) 
    FindText = "informacion de retiro" + "|" + "información de retiro" + "|" + "notificación retiro de personal"
    isRetired = re.findall(FindText, msg.Subject.lower())
    if isRetired:
        return True
    else:
        return False

def BrowserTabs():
    """ Create New Tabs on browser.

    This function generates a new tab in the browser for each attached mail file and place the focus on it.

    Parameters

    pos -- The number of .msg files.
    """
    driver.execute_script("window.open()")
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])

def CreateTicket(pos):
    """ Create Ticket on Help application.

    Extract the content of the email message file (msg variable), connect to the Help application and fill out the fields requested for the creation of the ticket.

    Parameters

    pos -- The number of .msg files.
    """
    msg = outlook.OpenSharedItem(PurePath(filename[pos]))                       # Contains the content of the email (title, body, etc.)
    driver.get(SiteInfo["website"])                                             # Initialize the opening of the browser.
    driver.get(SiteInfo["newticket"])
    FindElementOnHelp(msg)                                                      # Write in text field on help app
    SelectElementOnHelp(msg)                                                    # Select ComboBox on help app
    AttachmentOnHelp(msg, pos)
    time.sleep(2)                                                               # This field takes time to load the information, 2 seconds of waiting are added before writing the subcategory.
    select_element = Select(driver.find_element_by_id(SiteInfo["h_subcategory"]))  # The subcategory associated with the request is selected.
    select_element.select_by_visible_text("Retiro empleado")

def FindElementOnHelp(msg):
    """ Find text field on the Helpdesk app

    Search and find each text field in the Helpdesk application to write the content of the email message.

    Parameters

    msg -- The email message file with its formatted content.
    """
    element = driver.find_element_by_name(SiteInfo["h_user"])
    element.send_keys("mplata")
    element.send_keys(Keys.TAB)
    element = driver.find_element_by_id(SiteInfo["h_tittle"])                   # The title of the ticket is written.
    element.send_keys(msg.Subject)
    element = driver.find_element_by_id(SiteInfo["h_body"])                     # The body of the email is written as the case description to the ticket.
    bodysplit = msg.Body.split()                                                # Delete line breaks from message body
    element.send_keys(' '.join(bodysplit))                                      # Formatting the body of the message to avoid typing errors in the help application.

def SelectElementOnHelp(msg):
    """ Find ComboBox option on the Helpdesk app

    Search and find each ComboBox Option in the Helpdesk application to write the content according the email message.

    Parameters

    msg -- The email message file with its formatted content.
    """
    select_element = Select(driver.find_element_by_id(SiteInfo["h_category"]))  # The category associated with the request is selected.
    select_element.select_by_visible_text("Solicitudes")
    select_element = Select(driver.find_element_by_id(SiteInfo["h_priority"]))  # Case priority is selected in the ticket (default is low).
    select_element.select_by_visible_text("Baja")
    select_element = Select(driver.find_element_by_id(SiteInfo["h_contact"]))   # The user's contact mode is selected (default is Email).
    select_element.select_by_visible_text("Email")
    select_element = Select(driver.find_element_by_id(SiteInfo["h_ubication"])) # The user's location is selected.
    select_element.select_by_visible_text("CDJ")
    select_element = Select(driver.find_element_by_id(SiteInfo["h_department"])) # The area or department of the user is selected.
    select_element.select_by_visible_text("GESTION HUMANA") 

def AttachmentOnHelp(msg, pos):
    """ Attach mail file to the ticket

    Search and find attach button in the Helpdesk application to upload the email message file.

    Parameters

    msg -- The email message file with its formatted content.

    pos -- The number of .msg files.
    """
    button = driver.find_element_by_xpath(SiteInfo["h_idbutton"])               # find the "Adjuntar" button
    ActionChains(driver).click(button).perform()
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])  # Manipulate the new windows for attachment message file
    element = driver.find_element_by_xpath(SiteInfo["h_idfile"])
    element.send_keys(filename[pos])                                            # The file path is written to attach the email to the ticket.
    button = driver.find_element_by_xpath(SiteInfo["h_idsend"])
    ActionChains(driver).click(button).perform()
    driver.switch_to.alert.accept()
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])

for file in range(len(filename)):
    if IsRetiredUser(file):
        for i in range(7):
            CreateTicket(file)
            if (i+1) < 7:
                BrowserTabs()
        if (file + 1) <= len(filename):
            BrowserTabs()
    else:
        CreateTicket(file)
        if (file + 1) < len(filename):
            BrowserTabs()