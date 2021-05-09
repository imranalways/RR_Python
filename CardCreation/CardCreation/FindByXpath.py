from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select


def by_xpath(name_xpath, value_xpath, driver):
    if value_xpath is None or value_xpath.strip() == '':
        pass
    else:
        driver.find_element_by_xpath(name_xpath).clear()
        button_element = WebDriverWait(driver, 1).until(
            lambda chromedriver: driver.find_element_by_xpath(name_xpath))
        button_element.send_keys(value_xpath.strip())


def click_xpath(button_xpath, driver):
    WebDriverWait(driver, 1).until(lambda chromedriver: driver.find_element_by_xpath(button_xpath)).click()


def select_dropdown_text(select_xpath, select_value, driver):
    if select_value is None or select_value.strip() == '' or select_value == 'NS':
        pass
    else:
        select = Select(driver.find_element_by_xpath(select_xpath))
        select.select_by_visible_text(select_value)


def select_dropdown_by_value(select_xpath, select_value, driver):
    if select_value is None or select_value.strip() == '' or select_value == 'NS':
        pass
    else:
        select = Select(driver.find_element_by_xpath(select_xpath))
        select.select_by_value(select_value)


def select_dropdown_element(select_element, select_value, driver):
    if select_value is None:
        pass
    else:
        select = WebDriverWait(driver, 1).until(lambda chromedriver: driver.find_element_by_name(select_element))
        select.send_keys(select_value)
