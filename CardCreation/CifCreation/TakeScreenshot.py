import datetime
import time
import pyautogui
import ErrorHandler
import ConfigurationManager


def ScreenShot(startTime, loanId):
    FilePath = ConfigurationManager.RobotData("ImageFilePath")
    one_sec = int(ConfigurationManager.RobotData("DelayOneSec"))
    two_sec = int(ConfigurationManager.RobotData("DelayTwoSec"))
    screenShotTaken=True

    print("Start Calculating Coordinate---" + datetime.datetime.now().strftime("%H:%M:%S.%f"))
    x = None
    y = None

    #time.sleep(one_sec)

    while x is None:
        try:
            x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "Sign.PNG")
            #time.sleep(one_sec)
            print(x, y)
            Coordinate_X = x
            Coordinate_Y = y - 37
            print(Coordinate_X,Coordinate_Y)
            ###### Take ScreenShot:
            time.sleep(one_sec)
            im = pyautogui.screenshot(imageFilename=(FilePath.strip() + "SQDE_code.PNG"),
                                      region=(Coordinate_X, Coordinate_Y, 400, 50))
            time.sleep(one_sec)
            screenShotTaken = True

        except:
            x = None
            errorOccured = ErrorHandler.findError(startTime, loanId)
            if errorOccured is True:
                screenShotTaken = False
                break

    return screenShotTaken
