import pyautogui
import time
import datetime
import TakeScreenshot
import ImageToText
import ErrorHandler
import Calling_error_SP
import DatabaseManager
import ConfigurationManager


def entry():
    FilePath = ConfigurationManager.RobotData("ImageFilePath")
    common_Waiting_Time =float(ConfigurationManager.RobotData("CommonWaitingTime"))
    one_sec = int(ConfigurationManager.RobotData("DelayOneSec"))
    two_sec =int(ConfigurationManager.RobotData("DelayTwoSec"))
    count = 1
    noDataFoundLoopNo = 0
    while True:
        startTime = time.time()
        # region -------------------Fetching data from database--------------------------------------------
        print("Collecting Data...")
        results = DatabaseManager.ExecuteStoredProcedure("RDMS_GetCIFCreationData_Card_Rpa_py")
        # endregion

        if results.__len__() == 1:
            print("Processing...")
            noDataFoundLoopNo = 0
            # region -------------------------------Main screen image matching---------------------------------
            location = None
            time.sleep(two_sec)
            while location is None:
                try:
                    location = pyautogui.locateOnScreen(FilePath.strip() + "Location.PNG")
                except:
                    location = None
                    print("Main screen image does not match")
                    #return;
            # endregion

            for each_row in results:  # iterating the results array

                Loan_ID = each_row['loanId']
                print("loanId: ", Loan_ID)
                #clear form
                pyautogui.press('esc')
                pyautogui.press('esc')

                #region------------------------------- Date of Birth: ----------------------------------------------
                Dob_of_each_customer = each_row['DOB']
                pyautogui.typewrite(Dob_of_each_customer)
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                # endregion

                #region------------------------------- Customer Id: ----------------------------------------------
                #pyautogui.typewrite("")
                # time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                # endregion
                time.sleep(common_Waiting_Time)
                #region------------------------------- Title: ----------------------------------------------
                Title_of_each_customer = each_row['PRDTitle']
                pyautogui.typewrite(Title_of_each_customer.strip())
                pyautogui.press('tab')
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion
                # region------------------------------- First Name : ----------------------------------------------
                time.sleep(one_sec)
                FirstName_of_each_customer = each_row['FirstName']
                pyautogui.typewrite(FirstName_of_each_customer.strip())

                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion
                # region------------------------------- Middle Name : ----------------------------------------------
                MiddleName_of_each_customer = ""
                pyautogui.typewrite(MiddleName_of_each_customer.strip())

                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Last Name : ----------------------------------------------
                LastName_of_each_customer = each_row['LastName']
                pyautogui.typewrite(LastName_of_each_customer.strip())

                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Document Type : ----------------------------------------------
                DocType_of_each_customer = each_row['DocType']
                pyautogui.typewrite(DocType_of_each_customer.strip())
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Document Code : ----------------------------------------------
                DocCode_of_each_customer = each_row['DocCode']
                pyautogui.typewrite(DocCode_of_each_customer.strip())
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Unique ID : ----------------------------------------------
                UniqueId_of_each_customer = each_row['UniqueId']
                pyautogui.typewrite(UniqueId_of_each_customer.strip())

                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Address Line 1: ----------------------------------------------
                AddressLine1_of_each_customer = each_row['PRDAddress1']
                time.sleep(common_Waiting_Time)
                pyautogui.typewrite(AddressLine1_of_each_customer.strip())
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Address Line 2: ----------------------------------------------
                AddressLine2_of_each_customer = each_row['PRDAddress2']
                if AddressLine2_of_each_customer == None or AddressLine2_of_each_customer == '':
                    pyautogui.press('tab')
                    time.sleep(common_Waiting_Time)
                else:
                    pyautogui.typewrite(AddressLine2_of_each_customer.strip())
                    pyautogui.press('tab')
                    time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Address Line 3: ----------------------------------------------
                AddressLine3_of_each_customer = each_row['PRDAddress3']
                if AddressLine3_of_each_customer == None or AddressLine3_of_each_customer == '':
                    pyautogui.press('tab')
                    time.sleep(common_Waiting_Time)
                else:
                    pyautogui.typewrite(AddressLine3_of_each_customer.strip())
                    pyautogui.press('tab')
                    time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region----------------------------------- City : ---------------------------------------------------
                City_of_each_customer = each_row['City']
                pyautogui.typewrite(City_of_each_customer.strip())
                pyautogui.press('tab')
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region----------------------------------- State : ---------------------------------------------------
                State_of_each_customer = each_row['State']
                pyautogui.typewrite(State_of_each_customer.strip())

                pyautogui.press('tab')
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region----------------------------------- Country : -------------------------------------------------
                Country_of_each_customer = each_row['PRDCountry']
                time.sleep(one_sec)
                pyautogui.press('tab')
                pyautogui.typewrite(Country_of_each_customer.strip())
                pyautogui.press('tab')
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Postal Code: ----------------------------------------------
                PostalCode_of_each_customer = each_row['PostalCode']
                pyautogui.typewrite(PostalCode_of_each_customer.strip())

                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Address Valid From: ----------------------------------------------
                AddressValidFrom_of_each_customer = datetime.datetime.now().strftime("%d/%m/%Y")
                pyautogui.typewrite(AddressValidFrom_of_each_customer.strip())
                pyautogui.press('tab')
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Phone Type: ----------------------------------------------
                pyautogui.typewrite("MOBILE PHONE 1")
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Phone No: ----------------------------------------------
                pyautogui.typewrite("+88")
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                PhoneNo_of_each_customer = each_row['PRDMobile']
                pyautogui.typewrite(PhoneNo_of_each_customer.strip())

                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Email Type: ----------------------------------------------
                pyautogui.typewrite("COMMUNICATION")
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Email ID: ----------------------------------------------
                PRDEmail_of_each_customer = each_row['PRDEmail']
                pyautogui.typewrite(PRDEmail_of_each_customer.strip())

                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Nationality: ----------------------------------------------
                Nationality_of_each_customer = each_row['Nationality']
                pyautogui.press('tab')
                time.sleep(one_sec)
                pyautogui.typewrite(Nationality_of_each_customer.strip())
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Non - resident Indicator: -------------------------------------
                NonResidentIndicator_of_each_customer = "N"
                pyautogui.typewrite(NonResidentIndicator_of_each_customer.strip())
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Occupation: -------------------------------------
                Occupation_of_each_customer = each_row['occupation']
                if Occupation_of_each_customer == None or Occupation_of_each_customer == '':
                    pyautogui.press('tab')
                    time.sleep(common_Waiting_Time)
                else:
                    pyautogui.typewrite(Occupation_of_each_customer.strip())
                    pyautogui.press('tab')
                    time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Employment Type: -------------------------------------
                Employment_type_of_each_customer = each_row['EmploymentType']
                if Employment_type_of_each_customer == None or Employment_type_of_each_customer == '':
                    pyautogui.press('tab')
                    time.sleep(common_Waiting_Time)
                else:
                    pyautogui.typewrite(Employment_type_of_each_customer.strip())
                    pyautogui.press('tab')
                    time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Preferred Language: -------------------------------------
                PreferredLanguage_of_each_customer = "ENGLISH"
                if PreferredLanguage_of_each_customer == None or PreferredLanguage_of_each_customer == '':
                    pyautogui.press('tab')
                    time.sleep(common_Waiting_Time)
                else:
                    pyautogui.typewrite(PreferredLanguage_of_each_customer.strip())
                    pyautogui.press('tab')
                    time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- SOL : -------------------------------------
                PrimarySOLID_of_each_customer = each_row['SOLID']
                time.sleep(common_Waiting_Time)
                pyautogui.typewrite(PrimarySOLID_of_each_customer.strip())
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(one_sec)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                pyautogui.press('tab')
                time.sleep(common_Waiting_Time)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------------- Submit Button Click : -------------------------------------
                x = None
                y = None
                while x == None:
                    try:
                        x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "Submit.PNG")
                    except:
                        x = None
                        errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                        if errorOccured is True:
                            break
                pyautogui.click(x, y)
                errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                if errorOccured is True:
                    break
                # endregion

                # region------------------------- Take screen shoot for reading CIF : --------------------------------
                screenShotTaken = TakeScreenshot.ScreenShot(startTime, Loan_ID)
                if screenShotTaken is False:
                    break

                Cif = ImageToText.ReadText()
                if Cif == "Empty":
                    Calling_error_SP.error_sp(Loan_ID)
                    pyautogui.press('esc')
                    x,y = pyautogui.locateCenterOnScreen(FilePath.strip() + "close.PNG")
                    pyautogui.click(x,y)
                    pyautogui.press('enter')
                    break

                params = {}
                params["loanId"] = Loan_ID
                params["Cif"] = Cif
                DatabaseManager.ExecuteNonQueryStoredProcedure("RDMS_UpdateCIFNo_Card_Rpa", params)
                # click ok after reading cif:
                pyautogui.press('enter')
                time.sleep(one_sec)
                # endregion

                # region ------------------------- CIF saving proccece: --------------------------------
                x = None
                y = None
                while x is None:
                    try:
                        x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "CIFsave.PNG")
                        time.sleep(one_sec)
                    except:
                        x = None
                        errorOccured = ErrorHandler.findError(startTime, Loan_ID)
                        if errorOccured is True:
                            break
                pyautogui.click(x, y)
                pyautogui.typewrite("CIFCustomerApprovalEdit-10.2")
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                time.sleep(one_sec)
                pyautogui.press('enter')

                time.sleep(one_sec)
                pyautogui.press('enter')
                count = count + 1
                # endregion

        else:
            print("No data found. Waiting for data...")
            if noDataFoundLoopNo < 5:
                noDataFoundLoopNo = noDataFoundLoopNo + 1
                time.sleep(20)
            elif noDataFoundLoopNo >= 5:
                noDataFoundLoopNo = 0
                x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "MenuButton.PNG")
                pyautogui.click(x, y)
                time.sleep(one_sec)
                x, y = pyautogui.locateCenterOnScreen(FilePath.strip() + "Menu-4.PNG")
                pyautogui.click(x, y)
                time.sleep(one_sec)

            # break
