# https://www.selenium.dev/documentation/es/webdriver/browser_manipulation/
# https://selenium-python.readthedocs.io/installation.html
# https://docs.hektorprofe.net/python/funcionalidades-avanzadas/expresiones-regulares/

import win32com.client                                                          # import library to manipulate .msg files                                                                # Graphic library for interface
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
ConfigFile.read(r"C:\Recursos\Programacion\Python\Python_exercises\config.ini")
UserInfo = ConfigFile["USERINFO"]
SiteInfo = ConfigFile["HELP"]
filename = askopenfilenames(title = "Select outlook message file")              # Launch window to select outlook message file(s).
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")  # Object for manipulate outlook files
filedriver = UserInfo["chromewebdriverpath"]                                       # Path for chromedriver file to run browser.
options = Options()                                                             # Create option object for modify argument to chrome driver 
options.add_argument("--start-maximized")                                       # Maximized windows to start chrome driver
options.binary_location = UserInfo["chromeBinaryPath"] 
driver=webdriver.Chrome(options=options, executable_path=filedriver)            # Initialize the chrome browser 

def remove_urls (urltext):
    """ Check if the retired email contain url (<https://sap....) elements

    Check if the retired email contain url (<https://sap....) inside the mail and remove all elements.

    Parameters

    urltext -- string that body email message
    """

    urltext = re.sub(r'(<https|<http)?:\/\/(\w|\.|\/|\?|\=|\&|\%)*\b\>', '', urltext, flags=re.MULTILINE)
    return(urltext)

def IsRetiredUser(msg):
    """ Check if it is a user retired email

    Search the message title for the words associated with a user retired email.

    Parameters

    msg -- email message in the tuple class
    """

    FindText = "informacion retiro" + "|" + "información retiro" + "|" + "informacion de retiro" + "|" + "información de retiro" + "|" + "notificación retiro de personal" + "|" + "notificación retiro de personal" + "|" + "notificación retiro" "|" + "notificacion retiro"
    isRetired = re.findall(FindText, msg.Subject.lower())
    if isRetired:
        return True
    else:
        return False

def BrowserTabs():
    """ Create New Tabs on browser.

    This function generates a new tab in the browser for each attached mail file and place the focus on it.
    """
    driver.execute_script("window.open()")
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])

def CreateTicket(pos, msg):
    """ Create Ticket on Help application.

    Extract the content of the email message file (msg variable), connect to the Help application and fill out the fields requested for the creation of the ticket.

    Parameters

    pos -- The number of .msg files.

    msg -- email message in the tuple class
    """
    driver.get(SiteInfo["website"])                                             # Initialize the opening of the browser.
    driver.get(SiteInfo["newticket"])
    FindElementOnHelp(msg)                                                      # Write in text field on help app
    SelectElementOnHelp(msg)                                                    # Select ComboBox on help app
    AttachmentOnHelp(msg, pos)
    

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
    time.sleep(2)                                                               # This field takes time to load the information, 2 seconds of waiting are added before writing the subcategory.
    select_element = Select(driver.find_element_by_id(SiteInfo["h_subcategory"]))  # The subcategory associated with the request is selected.
    select_element.select_by_visible_text("Retiro empleado")

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
    time.sleep(2)
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    
for file in range(len(filename)):
    msg = outlook.OpenSharedItem(PurePath(filename[file]))                       # Contains the content of the email (title, body, etc.)
    subject = msg.Subject
    if IsRetiredUser(msg):
        for i in range(4):
            if (i == 0):
                msg.Subject = "Dominio, caja menor, MDM, Pin de impresión - " + subject
            elif (i == 1):
                msg.Subject = "Activos Fijos, Celular - " + subject
            elif (i == 2):
                msg.Subject = "Aplicaciones SAP y No SAP - " + subject
            elif (i == 3):
                msg.Subject = "Licencia Teams, Prominer, Línea celular - " + subject
            """elif (i == 4):
                msg.Subject = "Pin de impresión - " + subject
            elif (i == 5):
                msg.Subject = "Activos Fijos - " + subject
            elif (i == 6):
                msg.Subject = "Línea Celular - " + subject
            elif (i == 7):
                msg.Subject = "Licencia Prominer - " + subject"""

            # Remove all <https://hcm19.sapsf.com/ui/uicore/img/... text from message body on retired message from GH+
            msg.Body = remove_urls(msg.Body)

            CreateTicket(file, msg)
            if (i+1) < 4:
                BrowserTabs()

        if (file + 1) < len(filename):
            BrowserTabs()
    else:
        CreateTicket(file, msg)
        if (file + 1) < len(filename):
            BrowserTabs()