import pyautogui

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import openpyxl
import time
import pyodbc
import configuration

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=HOFS2406106155A\SQLEXPRESS;'
                      'Database=eOperations;'
                      'UID=sa;'
                      'PWD=biTS1234;')


cursor = conn.cursor()
cursor.execute('SELECT * FROM eOperations.dbo.RR_ORMUsers')

driver = webdriver.Chrome('C:/Users/imran.hossain/Downloads/chromedriver_win32/chromedriver.exe')
driver.maximize_window()
# ====================================LOGIN STARTS========================================
username="imran.hossain"
# username="debashish.mondal"
driver.get("http://localhost:55194/")
queUser = driver.find_element_by_xpath("/html/body/form/div[3]/div[2]/div/table/tbody/tr/td[1]/div/div/div/table/tbody/tr[4]/td/input")
time.sleep(2)
queUser.send_keys(username)
time.sleep(2)
quePass = driver.find_element_by_xpath("/html/body/form/div[3]/div[2]/div/table/tbody/tr/td[1]/div/div/div/table/tbody/tr[6]/td/input")
quePass.send_keys("abc")
queLogin = driver.find_element_by_xpath("/html/body/form/div[3]/div[2]/div/table/tbody/tr/td[1]/div/div/div/table/tbody/tr[8]/td/input")
queLogin.send_keys(Keys.RETURN)
# ====================================LOGIN ENDS========================================
time.sleep(3)

# ====================================NEW RISK CREATE STARTS========================================
xlDataFile = configuration.fileDataxlsx()
xlFile = configuration.fileXpathxlsx()

if username == "imran.hossain":
    newRisk=driver.find_element_by_xpath("/html/body/form/div[3]/div[2]/div/ul/li[2]/a")
    newRisk.send_keys(Keys.RETURN)
    time.sleep(2)
    i = 0

    for sl in xlDataFile.SL:
        print("hello:" + str(sl))
        for item in xlFile.Xpath:
            riskField = driver.find_element_by_xpath(item)
            time.sleep(1)

            if xlFile.type[i] == "click":
                riskField.send_keys(Keys.RETURN)

            elif xlFile.type[i] == "date":
                date = xlDataFile['' + xlFile.FieldTitle[i] + ''][sl-1].strftime("%d/%m/%Y")
                riskField.send_keys(date)
                lbl=driver.find_element_by_xpath("/html/body/form/div[3]/div[3]/div[2]/div/div[1]/div/div[2]/div/div[12]/div/input")
                lbl.send_keys(Keys.ENTER)

            else:
                text = xlDataFile['' + xlFile.FieldTitle[i] + ''][sl-1]
                riskField.send_keys(text)

            i = i+1
        i = 0
        time.sleep(1)
    time.sleep(1)
    # ====================================NEW RISK CREATE ENDS========================================
    # ====================================ORM RISK CHECKER STARTS========================================
    # print(driver.current_url)
    # if driver.current_url == "http://localhost:55194/UI/ORMList":


    # driver.get("http://localhost:55194/UI/ORMList")
    # time.sleep(1)
    # confirm=pyautogui.confirm('Do you  want to delete all risk?')  # returns "OK" or "Cancel"
    # if confirm == "OK":
    #     cursor.execute('EXEC dbo.RR_ORMList_Get')
    #     for item in cursor:
    #         dlt=driver.find_element_by_xpath("/html/body/form/div[3]/div[3]/div[2]/div[2]/div/div/div/div/div[2]/div/div/div[2]/table/tbody/tr[1]/td[1]/a[3]")
    #         dlt.send_keys(Keys.RETURN)
    #         time.sleep(1)
    #         dltConfirm=driver.find_element_by_xpath("/html/body/form/div[3]/div[3]/div[3]/div/div/div[3]/input")
    #         dltConfirm.send_keys(Keys.RETURN)
    #         time.sleep(1)

# ====================================ORM RISK CHECKER ENDS========================================

# elif username == "debashish.mondal":
#     time.sleep(2)
#     driver.get("http://localhost:55194/UI/ORMList")
#     time.sleep(2)
#
#     edt = driver.find_element_by_xpath("/html/body/form/div[3]/div[3]/div[2]/div[2]/div/div/div/div/div[2]/div/div/div[2]/table/tbody/tr[1]/td[1]/a[1]")
#     time.sleep(1)
#     edt.send_keys(Keys.RETURN)
#     time.sleep(2)
#
#     riskTitle=driver.find_element_by_xpath("/html/body/form/div[3]/div[3]/div[2]/div/div[1]/div[1]/div[1]/div/div[2]/div/input")
#     print(riskTitle.get_attribute("value"))
#     wb = openpyxl.Workbook()
#     sheet = wb.active
#     row=1
#     column=1
#     for serial in xlDataFile.SL:
#         for item in xlDataFile.columns:
#             c1 = sheet.cell(row, column)
#             c1.value = xlDataFile.columns[column-1]
#             column = column+1
#         row=row+1
#         column = 1
#     wb.save(r'C:\Users\imran.hossain\Downloads\ALL Files\PYTHON\Risk_Resister_Python\editData.xlsx')


conn.commit()


