from selenium.webdriver.support.wait import WebDriverWait
import time


class login(object):

    def tillCardCreate(self,driver):

        cardproUsername = 'sumon27257'
        cardproPassword = 'sumon@27'
        userFieldname = 'user_id'
        userPasswordname = 'passwd'
        submitButtonname = 'submit'
        continueButtonname = 'Continue'

        userfieldElement = WebDriverWait(driver, 1).until(lambda driver: driver.find_element_by_name(userFieldname))
        userPasswordElement = WebDriverWait(driver, 1).until(
            lambda driver: driver.find_element_by_name(userPasswordname))
        submitButtonElement = WebDriverWait(driver, 1).until(
            lambda driver: driver.find_element_by_name(submitButtonname))
        userfieldElement.send_keys(cardproUsername)
        userPasswordElement.send_keys(cardproPassword)
        time.sleep(2)
        submitButtonElement.click()
        time.sleep(2)

        try:
            click_buttonXpath = '/html/body/table[2]/tbody/tr/td/form/table/tbody/tr/td/input[1]'
            WebDriverWait(driver, 1).until(lambda driver: driver.find_element_by_xpath(click_buttonXpath)).click()
            continueButtonElement = WebDriverWait(driver, 1).until(
                lambda driver: driver.find_element_by_name(continueButtonname))
            continueButtonElement.click()
        except:
            continueButtonElement = WebDriverWait(driver, 1).until(
                lambda driver: driver.find_element_by_name(continueButtonname))
            continueButtonElement.click()
        time.sleep(3)

        APM_buttonXpath = '/html/body/table[3]/tbody/tr/td[4]/div/table/tbody/tr[2]/td[2]/a'
        WebDriverWait(driver, 1).until(lambda driver: driver.find_element_by_xpath(APM_buttonXpath)).click()
        NADC_buttonXpath = '/html/body/table[3]/tbody/tr/td[4]/div/table/tbody/tr[2]/td[2]/a'
        WebDriverWait(driver, 1).until(lambda driver: driver.find_element_by_xpath(NADC_buttonXpath)).click()
        clickradio_buttonXpath = '/html/body/table[3]/tbody/tr/td[4]/form/table[2]/tbody/tr[3]/td[1]/input[1]'
        WebDriverWait(driver, 1).until(lambda driver: driver.find_element_by_xpath(clickradio_buttonXpath)).click()
        update1_buttonXpath = '/html/body/table[3]/tbody/tr/td[4]/form/table[3]/tbody/tr/td/input[2]'
        WebDriverWait(driver, 1).until(lambda driver: driver.find_element_by_xpath(update1_buttonXpath)).click()
        time.sleep(1)
        driver.switch_to.alert.accept()
        time.sleep(1)