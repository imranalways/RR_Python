import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

import ConfigurationManager
import EncryptDecrypt
from BasicCard import basic_card
from Connection import databaseConnection
import logging

ServerName = EncryptDecrypt.Decrypt(ConfigurationManager.ConnectionString("Host"))  #3NA5yj9jSmFFxQRbH4gbiV8ChksQQwGW
Database = EncryptDecrypt.Decrypt(ConfigurationManager.ConnectionString("ServiceName")) #hQbccBHsqeuDUDiAm4mFlw==
UserId = EncryptDecrypt.Decrypt(ConfigurationManager.ConnectionString("UserName")) #aUTWiFily1872EzTSbNzCdtnlK71FOXcp093EWnk/Qk=
Password = EncryptDecrypt.Decrypt(ConfigurationManager.ConnectionString("Password")) #AbheEDfv95LpSvjSKTnxcg==

baseURL = ConfigurationManager.RobotData("HomeURL") #http://172.25.4.222:7001/bosprod/
driverLocation = ConfigurationManager.RobotData("DriverLocation") #C:\RPA_Cardpro\Card_Creation\chromedriver.exe
# driverLocation = ConfigurationManager.RobotData("FireFoxDriverLocation")
WaitTimeForBrowserRefresh = int(ConfigurationManager.RobotData("WaitTimeForBrowserRefresh")) #15secends
# FilePath = ConfigurationManager.RobotData("ImageFilePath")

logging.basicConfig(filename='CardCreation.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')


# connectionString = 'DRIVER={SQL Server};SERVER=' + ServerName.strip() + ';DATABASE=' + Database.strip() \
#                    + ';UID=' + UserId.strip() + ';PWD=' + Password.strip() + ';'

connectionString = 'DRIVER={SQL Server};SERVER=HOFS2406106155A\SQLEXPRESS;' \
                   'DATABASE= eOperations; UID=sa;PWD=biTS1234;'

option = webdriver.ChromeOptions()
option.add_experimental_option("useAutomationExtension", False)
option.set_capability('unhandledPromptBehavior', 'ignore')
chromeCapabilities = DesiredCapabilities.CHROME.copy()
chromeCapabilities['marionette'] = False
driver = webdriver.Chrome(desired_capabilities=chromeCapabilities,
                          executable_path=driverLocation,
                          options=option)
driver.get(baseURL)
driver.maximize_window()
driver = driver


def until_find():
    try:
        WebDriverWait(driver, 7).until(
            ec.presence_of_element_located((By.XPATH, "/html/body/table[3]/tbody/tr/td[3]/table/tbody/tr/td[2]"))
        )
        return True
    except Exception as sample:
        logging.warning('_From-> until_find()_ :' + str(sample))
        return False


def accept_alert():
    try:
        WebDriverWait(driver, 7).until(ec.alert_is_present(),
                                       'Timed out waiting for alerts to appear in accept_alert')
        driver.switch_to.alert.accept()
        return 'accepted'
    except Exception as sample:
        logging.warning('_From-> accept_alert()_ :' + str(sample))
        return Exception


try:
    isTrue = False
    while isTrue is False:
        try:
            isTrue = until_find()
            print(isTrue)
        except Exception:
            isTrue = False

    procedureName = "{call RDMS_GetCreditCardInfoBySerialNo_Card_Rpa_py}"
    objectConnection = databaseConnection()
    results = objectConnection.callProcedure(procedureName, connectionString)

    while True:
        while results.__len__() == 1:
            try:
                basic_card(driver, results, connectionString, logging)

                results = []
                procedureName = "{call RDMS_GetCreditCardInfoBySerialNo_Card_Rpa_py}"
                objectConnection = databaseConnection()
                results = objectConnection.callProcedure(procedureName, connectionString)

            except Exception as e:
                logging.warning('_From-> CardCreation()_ :' + str(e))
                accepted = accept_alert()

                try:
                    add1_buttonXpath = '/html/body/table[3]/tbody/tr/td[3]/form/table[2]/tbody/tr/td/input[3]'
                    WebDriverWait(driver, 1).until(
                        lambda chromedriver: driver.find_element_by_xpath(add1_buttonXpath)).click()
                    time.sleep(1)

                except Exception as e:
                    logging.warning('_From-> CardCreation->on returning main page_ :' + str(e))
                    add1_buttonXpath = '/html/body/table[3]/tbody/tr/td[3]/form/table[2]/tbody/tr/td/input[2]'
                    WebDriverWait(driver, 1).until(
                        lambda chromedriver: driver.find_element_by_xpath(add1_buttonXpath)).click()
                    time.sleep(1)

                results = []
                procedureName = "{call RDMS_GetCreditCardInfoBySerialNo_Card_Rpa_py}"
                objectConnection = databaseConnection()
                results = objectConnection.callProcedure(procedureName, connectionString)

        time.sleep(WaitTimeForBrowserRefresh)
        results = []
        procedureName = "{call RDMS_GetCreditCardInfoBySerialNo_Card_Rpa_py}"
        objectConnection = databaseConnection()
        results = objectConnection.callProcedure(procedureName, connectionString)

        print('Waiting For Data...')
        buttonUpdateBatchXpath = '/html/body/table[3]/tbody/tr/td[3]/form/table[4]/tbody/tr/td/input[4]'
        WebDriverWait(driver, 1).until(lambda chromedriver:
                                       driver.find_element_by_xpath(buttonUpdateBatchXpath)).click()
        driver.switch_to.alert.dismiss()
    print('success')

except Exception as e:
    print("Something May have gone wrong\nContact with the IT services OR"
          "\nPlease take to batch record lists & run again \n" + str(e))

finally:
    print("done")
    driver.quit()
