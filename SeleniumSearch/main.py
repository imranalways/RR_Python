from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import time

driver = webdriver.Chrome('C:/Users/imran.hossain/Downloads/chromedriver_win32/chromedriver.exe')
driver.get("http://www.google.com")




que=driver.find_element_by_xpath("//input[@name='q']")
que.send_keys("Software testing")
time.sleep(2)
que.send_keys(Keys.ARROW_DOWN)
que.send_keys(Keys.ARROW_DOWN)
time.sleep(2)
que.send_keys(Keys.RETURN)
