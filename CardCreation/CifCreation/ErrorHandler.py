import pyautogui
import requests
import io
import pytesseract
# import yaml
from PIL import Image
from datetime import timedelta
import time
import datetime
from time import sleep
import Calling_error_SP
import ConfigurationManager
import DatabaseManager


def findError(startTime, loanId):
    FilePath = ConfigurationManager.RobotData("ImageFilePath")
    error_wait_time = float(ConfigurationManager.RobotData("MaxWaitingTimeForError"))
    common_wait_time = float(ConfigurationManager.RobotData("CommonWaitingTime"))
    one_sec = int(ConfigurationManager.RobotData("DelayOneSec"))
    curTime = time.time()
    timeDiff = curTime - startTime
    compare = timedelta(seconds=round(timeDiff)).seconds
    if compare > error_wait_time:
        pyautogui.press('esc')
        x,y = pyautogui.locateCenterOnScreen(FilePath.strip() + "close.PNG")
        pyautogui.click(x,y)
        time.sleep(one_sec)
        location = None
        try:
            location = pyautogui.locateCenterOnScreen(FilePath.strip() + "ProcessSelection.PNG")
        except:
            pass
        if location != None:
            pyautogui.click(location)
            time.sleep(common_wait_time)
            pyautogui.press('tab')
            time.sleep(common_wait_time)
            pyautogui.press('tab')
            time.sleep(common_wait_time)
            pyautogui.press('tab')
            time.sleep(common_wait_time)
            pyautogui.press('tab')
            time.sleep(common_wait_time)
            pyautogui.press('tab')
            time.sleep(common_wait_time)
            pyautogui.press('tab')
            time.sleep(common_wait_time)
            pyautogui.press('tab')
            time.sleep(common_wait_time)
            pyautogui.press('tab')
            time.sleep(common_wait_time)
            # pyautogui.press('tab')
            # time.sleep(common_wait_time)
            pyautogui.press('enter')
            time.sleep(common_wait_time)
            pyautogui.press('enter')
            time.sleep(common_wait_time)

        pyautogui.press('enter')

        # error status update
        Calling_error_SP.error_sp(loanId)

        return True

    return False
