import pyautogui
import time
import DatabaseManager
from MiscellaneousTask import AdditionalTask


class AccOpen(object):

    def entry(self, data, filepath, page_name, logger):
        params = {"PageName": page_name}
        attribute = DatabaseManager.ExecuteStoredProcedureWithParam("RPA_CIF_QDE_Get_Attribute", params)

        for item in attribute:

            if item['AttributeType'] == 'BUTTON':

                # region------------------------------- Find Button and click only : -----------------------------------
                xy = None
                while xy is None:
                    try:
                        xy = pyautogui.locateCenterOnScreen(filepath.strip() + item['ImageName'])

                    except Exception as e:
                        logger.error(e.__str__() + ' ' + item['ImageName'])
                        xy = None

                if item['HasMultipleEntry'] is True:
                    additional_task = AdditionalTask()
                    additional_task.do_multiple_task(filepath, item['AttributeLabel'], data[0]['GuarantorId'], logger)
                else:
                    pyautogui.click(xy[0], xy[1])

                if item['IsLoadingDelay'] is True:
                    time.sleep(item['LoadingDelayTime'])

                if item['NeedToScrollDown'] is True:
                    pyautogui.scroll(-250)
                    time.sleep(.5)

                if item['NeedToScrollUp'] is True:
                    pyautogui.scroll(2500)
                    time.sleep(.5)
                    pyautogui.scroll(2500)
                    time.sleep(.5)

            else:
                if item['ImageName'] != '':
                    # region------------------------------- Find Menu : -------------------------------------
                    y = None
                    x = None
                    while x is None:
                        try:
                            x, y = pyautogui.locateCenterOnScreen(filepath.strip() + item['ImageName'])
                        except Exception as e:
                            x = None
                            logger.error(e.__str__() + ' ' + item['ImageName'])

                    pyautogui.click(x, y)
                    time.sleep(.5)
                    pyautogui.press('tab')

                    if item['WantToPressTab'] is True:
                        for i in range(item['TabPressCount']):
                            time.sleep(.5)
                            pyautogui.press('tab')

                    if item['WantToClearField'] is True:
                        pyautogui.press('backspace')

                    if item['IsFixedValue'] is True:
                        pyautogui.typewrite(item['FixedValue'])
                    else:
                        pyautogui.typewrite(data[0][item['DatabaseColumnAlias']])

                    if item['IsLoadingDelay'] is True:
                        pyautogui.press('tab')
                        time.sleep(item['LoadingDelayTime'])

                    if item['SleepingTime'] > 0:
                        sleeping_time = int(item['SleepingTime'])
                        time.sleep(sleeping_time)

                    if item['NeedToScrollDown'] is True:
                        pyautogui.scroll(-250)
                        time.sleep(.5)

                    if item['NeedToScrollUp'] is True:
                        pyautogui.scroll(1000)
                        time.sleep(.5)
