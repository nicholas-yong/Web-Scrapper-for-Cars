import xlwings as xw
from selenium import webdriver 
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC 
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException

import time
import random

#Constants
make = ["Toyota", "Honda", "Ford", "Kia", "Mazda"]
location = "Western%20Australia"
price_max = 10000
keywords = "Automatic"
max_odometer = 100000
proxy_list_url = "https://www.sslproxies.org/"

#comment this out for now, we currently doesn't use this method of selecting a vehicle
#def comboBoxSelection( browser, parentName, selectorValue ):
#    element = WebDriverWait(browser, timeout).until(EC.element_to_be_clickable((By.CSS_SELECTOR, parentName)))
#    actionChains = ActionChains(browser)
#    actionChains.click(element).perform()
    #Once the list has been generated...
#    child_element = WebDriverWait(browser, timeout).until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"{parentName} > .dropdown-box > div[title={make}]")))
#    actionChains_Two = ActionChains(browser)
#    actionChains_Two.click(child_element).perform()

def getProxyList(browser, timeout):
    browser.get(proxy_list_url)
    ip_list = WebDriverWait(browser, timeout).until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, f"#proxylisttable > tbody > tr[role=row] > td:nth-child(1)" )))
    port_list = WebDriverWait(browser, timeout).until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, f"#proxylisttable > tbody > tr[role=row] > td:nth-child(2)" )))
    proxy_list = []

    for i in range(len(ip_list)):
        #ip_list and port_list should have the same amount of elements.
        proxy_list.append( f"{ip_list[i].text}:{port_list[i].text}" ) 

    return proxy_list

def main():
    #Build the query required to reach the search results first.
    #CarSales.com
    #Get the Proxy_list first.
    #Initialization Code
    timeout = 30
    option = webdriver.ChromeOptions()
    option.add_argument( "-incongito")
    browser = webdriver.Chrome(executable_path="C:/Users/nickh/Downloads/Application Setup/chromedriver.exe", chrome_options=option)
    proxy_list = getProxyList( browser, timeout)
    random.seed() 
    proxy_selector = random.randint(0, 19)

    #Create a new browser each time we send a request to the carsales with a new proxy.
    option.add_argument(f"--proxy-server={proxy_list[proxy_selector]}")
    print(proxy_list[proxy_selector])
    #Get the new driver
    browser = webdriver.Chrome(executable_path="C:/Users/nickh/Downloads/Application Setup/chromedriver.exe", chrome_options=option)

    #Construct the required query with the proper query strings.
    make_query = "And." if len(make) == 1 else "Or."
    for index, item in enumerate(make):
        make_query = make_query + "Make." + item + ( "." if index == len(make) - 1 else "._." )

    query = f"(And.Service.CARSALES._.({make_query})_.State.{location}._.Price.range(..{price_max})._.CarAll.keyword({keywords})._.Odometer.range(..{max_odometer}).)"

    #Direct browser to the carsales website
    browser.get(f"https://www.carsales.com.au/cars/?&q=({query}")
    
    #Once the webpage has loaded, we need to set the options for the form value.
    #Wait till the page has loaded
    #timeout = 30
    #WebDriverWait(browser, timeout).until(EC.visibility_of_element_located((By.XPATH, "//div[@class='carsales']")))

    #Set form values
    #comboBoxSelection("#search-field-make", make )
    #comboBoxSelection("#search-field-location", location )
    #comboBoxSelection("#search-field-location", location )

    time.sleep(10000)
    



@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("CarsSheet.xlsm").set_mock_caller()
    main()

