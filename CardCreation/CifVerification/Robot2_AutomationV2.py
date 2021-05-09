import pyautogui
import time
import datetime
import ConfigurationManager
import DatabaseManager


def approve():

    noDataFoundLoopNo = 0

    MaxWaitingTimeForError = int(ConfigurationManager.RobotData("MaxWaitingTimeForError"))
    ImageExceptionWait = int(ConfigurationManager.RobotData("ImageExceptionWait"))
    FilePath = ConfigurationManager.RobotData("ImageFilePath")
    ApproveWait = ConfigurationManager.RobotData("ApproveWait")
    count = 1

    while True:
        print("collecting data")
        results = DatabaseManager.ExecuteStoredProcedure("RDMS_GetCifInfoForApproval_Card_Rpa")

        if results.__len__() == 1:
            noDataFoundLoopNo = 0

            location = None
            time.sleep(2)
            while location is None:
                try:
                    location = pyautogui.locateOnScreen(FilePath.strip() + "ui.PNG")

                except:
                    print("ui.png not found")
                    location = None

            # region -------------------------------Data entry:---------------------------------
            for each_row in results:

                start = time.time()

                loan_Id = each_row['loanId']  # taking each customers dob from the dict and puting it in a varaible

                time.sleep(1)
                pyautogui.typewrite("Business Center Group")
                time.sleep(1)
                pyautogui.press('tab')
                time.sleep(1)

                x = None
                y = None

                time.sleep(2)
                while x == None:
                    try:
                        x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "group.PNG")

                    except:
                        x = None
                        try:
                            x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "general.PNG")

                        except:
                            x = None
                pyautogui.click(x, y)

                time.sleep(1)
                pyautogui.typewrite("General Banking")
                time.sleep(1)

                CIF = each_row['CIF']  # taking each customers dob from the dict and puting it in a varaible

                Last_CIF = CIF

                time.sleep(1)
                pyautogui.press('tab')
                time.sleep(1)
                pyautogui.press('tab')
                time.sleep(1)
                pyautogui.typewrite(CIF)
                time.sleep(1)
                pyautogui.press('enter')
                time.sleep(2)
                pyautogui.press('enter')
                time.sleep(2)
                pyautogui.press('tab')
                pyautogui.press('enter')

                x = None
                y = None
                counter=0

                time.sleep(2)
                while x == None:
                    try:
                        x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "trybox.PNG")

                    except:
                        x = None
                        time.sleep(1)
                        counter +=1
                        if counter > ImageExceptionWait:
                            break
                pyautogui.click(x, y)

                x = None
                y = None
                counter = 0

                time.sleep(2)
                while x == None:
                    try:
                        x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "current process step.PNG")

                    except:
                        x = None
                        time.sleep(1)
                        counter += 1
                        if counter > ImageExceptionWait:
                            break
                pyautogui.click(x, y)

                time.sleep(2)

                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('enter')

                x = None
                y = None
                counter = 0

                time.sleep(int(ApproveWait))
                while x == None:
                    try:
                        x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "select approve.PNG")

                    except:
                        x = None
                        time.sleep(1)
                        counter += 1
                        if counter > ImageExceptionWait:
                            break

                pyautogui.click(x, y)

                x = None
                y = None

                time.sleep(int(ApproveWait))
                counter = 0
                while x == None:
                    try:
                        x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "Approve_selectwh.PNG")

                    except:
                        x = None
                        time.sleep(1)
                        counter += 1
                        if counter > ImageExceptionWait:
                            break
                pyautogui.click(x, y)

                # pyautogui.typewrite("Approve", 0.3)

                time.sleep(int(ApproveWait))
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('enter')

                x = None
                y = None
                counter = 0
                time.sleep(int(ApproveWait))
                while x == None:
                    try:
                        x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "ApprovalOK.PNG")
                    except:
                        x = None
                        time.sleep(1)
                        counter += 1
                        if counter > ImageExceptionWait:
                            break
                pyautogui.click(x, y)

                #time.sleep(2)
                #pyautogui.press('enter')

                done = time.time()
                elapsed = int(done - start)
                print(elapsed)
                # region ######################################## Update Approval Status ##############################################
                if elapsed < MaxWaitingTimeForError:

                    loanId = loan_Id

                    params = {}
                    params["loanId"] = loanId
                    DatabaseManager.ExecuteNonQueryStoredProcedure("RDMS_UpdateApprovedStatusCIFNo_Card_Rpa", params)

                    x = None
                    y = None

                    time.sleep(2)
                    while x == None:
                        try:
                            x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "down.PNG")
                        except:
                            x = None
                    pyautogui.click(x, y)

                    time.sleep(1)

                    x = None
                    y = None

                    time.sleep(2)
                    while x == None:
                        try:
                            x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "entity_id.PNG")
                        except:
                            x = None

                    X_coordinate = x + 287
                    Y_coordinate = y
                    pyautogui.click(X_coordinate, Y_coordinate)

                    pyautogui.press('backspace')
                    pyautogui.press('backspace')
                    pyautogui.press('backspace')
                    pyautogui.press('backspace')
                    pyautogui.press('backspace')
                    pyautogui.press('backspace')
                    pyautogui.press('backspace')
                    pyautogui.press('backspace')

                    x = None
                    y = None

                    time.sleep(2)
                    while x == None:
                        try:
                            x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "self.PNG")

                        except:
                            x = None
                    pyautogui.click(x, y)

                    time.sleep(2)
                    pyautogui.click(x, y)

                    count = count+1
                # endregion

                # region Error status #########################################################
                else:

                    loanId = loan_Id

                    params = {}
                    params["loanId"] = loanId
                    DatabaseManager.ExecuteNonQueryStoredProcedure("RDMS_UpdateCIFNoVerifyError_Card_Rpa", params)

                    x = None
                    y = None

                    time.sleep(2)
                    while x == None:
                        try:
                            x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "down.PNG")

                        except:
                            x = None
                    pyautogui.click(x, y)

                    time.sleep(1)

                    x = None
                    y = None

                    time.sleep(2)
                    while x == None:
                        try:
                            x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "entity_id.PNG")

                        except:
                            x = None

                    X_coordinate = x + 287
                    Y_coordinate = y
                    pyautogui.click(X_coordinate, Y_coordinate)

                    pyautogui.press('backspace')
                    pyautogui.press('backspace')
                    pyautogui.press('backspace')
                    pyautogui.press('backspace')
                    pyautogui.press('backspace')
                    pyautogui.press('backspace')
                    pyautogui.press('backspace')
                    pyautogui.press('backspace')

                    x = None
                    y = None

                    time.sleep(2)
                    while x == None:
                        try:
                            x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "for try.PNG")

                        except:
                            x = None
                    pyautogui.click(x, y)

                    time.sleep(2)
                    pyautogui.click(x, y)

                    count = count + 1
                # endregion

            # endregion
        else:
            if noDataFoundLoopNo < 5:
                noDataFoundLoopNo = noDataFoundLoopNo + 1
                time.sleep(13)
            elif noDataFoundLoopNo >= 5:
                print("Waiting for Approval data.......")
                noDataFoundLoopNo = 0
                x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "MenuButton.PNG")
                pyautogui.click(x, y)
                time.sleep(1)
                x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "EntityQueue.PNG")
                pyautogui.click(x, y)
                time.sleep(1)
