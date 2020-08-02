#https://www.selenium.dev/documentation/es/support_packages/working_with_select_elements/
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options

filename = r"C:\Recursos\Programacion\Python\Browser\chromedriver.exe"

from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains

c_option = Options() 
c_option.add_argument("--start-maximized") 
driver=webdriver.Chrome(ptions=c_option, executable_path=filename) 

user_name = "cmeneses"
tittle = "This is my test"
body = "This is the first connection from Selenium web driver"

#driver = webdriver.Chrome(options, filename)
driver.get("http://help")
driver.get("http://help/NuevoTicket.asp")
element = driver.find_element_by_name("usu_nombre")
element.send_keys(user_name)
element.send_keys(Keys.TAB)
element = driver.find_element_by_id("H_titulo")
element.send_keys(tittle)
element = driver.find_element_by_id("H_Descripcion")
element.send_keys(body)

#select_element = Select(driver.find_element_by_id("H_area"))
#select_object = Select(select_element)
#select_object.select_by_index(5)
#select_element.select_by_index(5)

#select_element = Select(driver.find_element_by_id("H_SubArea"))
#select_element.select_by_index(5)

select_element = Select(driver.find_element_by_id("H_prioridad"))
select_element.select_by_index(2)

select_element = Select(driver.find_element_by_id("Sel_Contacto"))
select_element.select_by_index(1)

select_element = Select(driver.find_element_by_id("H_Ubicacion"))
select_element.select_by_index(4)

select_element = Select(driver.find_element_by_id("H_OM_dpto"))
select_element.select_by_index(4)
print ("Esta es la ventana principal " + driver.current_window_handle)
button = driver.find_element_by_xpath("//*[@id='adjuntar_archivo']")
ActionChains(driver).click(button).perform()
driver.switch_to.window(driver.window_handles[1])
print ("Este es el cuadro de dialogo " + driver.current_window_handle)

element = driver.find_element_by_xpath("//*[@id='file']")
element.send_keys("C:\\test\\GH+.msg")

button = driver.find_element_by_xpath("//*[@id='enviar']")
ActionChains(driver).click(button).perform()

driver.switch_to.alert.accept()
print ("Cerrada la pestaña " + driver.current_window_handle)

driver.switch_to.window(driver.window_handles[0])

# Opens a new tab
driver.execute_script("window.open()")
# Switch to the newly opened tab
driver.switch_to.window(driver.window_handles[1])
print ("Nueva pestaña " + driver.current_window_handle)
driver.get("http://help")
driver.get("http://help/NuevoTicket.asp")

# Selecciona una <option> basándose en el indice interno del elemento <select>
#select_object.select_by_index(1)

# Selecciona una <option> basándose en su atributo value
#select_object.select_by_value('value1')

# Selecciona una <option> basándose en el texto que muestra
#select_object.select_by_visible_text('Bread')
  

#element.send_keys(Keys.RETURN)