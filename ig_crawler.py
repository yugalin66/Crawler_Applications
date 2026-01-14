from selenium import webdriver
from selenium.webdriver import ActionChains
import time
#再多import ActionChains

driver = webdriver.Chrome()
driver.get('https://www.instagram.com/reel/C6qYNfiyvuQ/')

for count in range(10000):   
    time.sleep(7)       
    driver.refresh()
