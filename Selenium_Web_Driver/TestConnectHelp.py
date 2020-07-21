#https://www.selenium.dev/documentation/es/support_packages/working_with_select_elements/
from selenium import webdriver
from selenium.webdriver.support.select import Select
filename = r"C:\Recursos\Programacion\Python\Browser\chromedriver.exe"

from selenium.webdriver.common.keys import Keys
user_name = "testuser"
tittle = "This is my test"
body = "This is the first connection from Selenium web driver"
driver = webdriver.Chrome(filename)
driver.get("http://help")
driver.get("http://help/NuevoTicket.asp")
element = driver.find_element_by_name("usu_nombre")
element.send_keys(user_name)
element.send_keys(Keys.TAB)
element = driver.find_element_by_id("H_titulo")
element.send_keys(tittle)
element = driver.find_element_by_id("H_Descripcion")
element.send_keys(body)

select_element = Select(driver.find_element_by_id("H_area"))
#select_object = Select(select_element)
#select_object.select_by_index(5)
select_element.select_by_index(5)

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

# Selecciona una <option> basándose en el indice interno del elemento <select>
#select_object.select_by_index(1)

# Selecciona una <option> basándose en su atributo value
#select_object.select_by_value('value1')

# Selecciona una <option> basándose en el texto que muestra
#select_object.select_by_visible_text('Bread')
  

#element.send_keys(Keys.RETURN)