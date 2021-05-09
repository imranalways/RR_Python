import webbrowser
import datetime

from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import time
import pyodbc

from Another import Another

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=HOFS2406106155A\SQLEXPRESS;'
                      'Database=TestDB;'
                      'UID=sa;'
                      'PWD=biTS1234;')

# Another()
cursor = conn.cursor()
cursor.execute('SELECT * FROM TestDB.dbo.Employees')

# cursor.execute('SELECT * FROM eOperations.dbo.RR_RiskMaster')
# cursor.execute("EXEC dbo.RR_Risk_GetAll")
# trackId="RMUImran1002"
# arg3 = ('INFO1101')
# cursor.callproc('EXEC dbo.RR_ORMUser_Insert', arg3)

# name="HASAN"
# cursor.execute('EXEC dbo.RR_ORMUser_Insert ? ,?, ?, ?, ?, ?, ?, ? ', name, "Pro", "6295", "imran@gmail.com", "73737388", "8282827738", 1, "Admin")
# cursor.execute("EXEC dbo.RR_ORMUser_GetAll")
# username = input("Enter username:")
# print(username)
trackId="RMUImran1002"
RiskTitle="Risk ABC Updated"
RiskType="THIS is Risk Type"
RiskTypeOther=""
RiskDescription="ThisisRiskdescription"
DateOfRiskDiscovery=datetime.datetime.now()
DateOfRiskReport=datetime.datetime.now()
RiskArea="NoRiskArea Is available"
ResponsibleDepartment="DEPT"
ResponsibleDepartmentOther=""
ResponsiblePerson="ABC"
IsRecurring=0
RecurringRiskDescription=""
HasOperationalLoss=0
OperationalLossDescription=""
TentativeRootCause="RES"
TentativeRootCauseOthers=""
BusinessImpact="AREES"
BusinessImpactOthers=""
IsActionTaken=0
ActionTakenDescription=""
ProvidersPin="4232"
Recommendation="hqsvd"
ProposedClosingDate=datetime.datetime.now()
Status="OKKstatus"
Segment="RET"
ReportedBy="imran"
xmla="<note><to>Tove</to><from>Jani</from><heading>Reminder</heading><body>Don't forget me this weekend!</body></note>"
# cursor.execute('EXEC dbo.RR_RiskRegister_Insert ? ,?, ?, ?, ?, ?, ?, ?,? ,?, ?, ?, ?, ?, ?, ?, ? ,?, ?, ?, ?, ?, ?, ?, ? ,?, ?, ?',
#                trackId, RiskTitle, RiskType, RiskTypeOther, RiskDescription, DateOfRiskDiscovery, DateOfRiskReport,
#                RiskArea, ResponsibleDepartment, ResponsibleDepartmentOther, ResponsiblePerson, IsRecurring,
#                RecurringRiskDescription, HasOperationalLoss, OperationalLossDescription,TentativeRootCause,
#                TentativeRootCauseOthers, BusinessImpact, BusinessImpactOthers, IsActionTaken, ActionTakenDescription,
#                ProvidersPin, Recommendation, ProposedClosingDate, Status, Segment, ReportedBy, xmla)


# cursor.execute('EXEC dbo.RR_RiskRegister_Update ? ,?, ?, ?, ?, ?, ?, ?,? ,?, ?, ?, ?, ?, ?, ?, ? ,?, ?, ?, ?, ?, ?, ?, ? ,?, ?, ?, ?, ?',
#                trackId, RiskTitle, RiskType, RiskTypeOther, RiskDescription, DateOfRiskDiscovery, DateOfRiskReport,
#                RiskArea, ResponsibleDepartment, ResponsibleDepartmentOther, ResponsiblePerson, IsRecurring,
#                RecurringRiskDescription, HasOperationalLoss, OperationalLossDescription,TentativeRootCause,
#                TentativeRootCauseOthers, BusinessImpact, BusinessImpactOthers, IsActionTaken, ActionTakenDescription,
#                ProvidersPin, Recommendation, ProposedClosingDate, Status, 0, ReportedBy, "6155", Segment,  xmla)


# chromedir = "C:/Program Files/Google/Chrome/Application/chrome.exe %s"
driver = webdriver.Chrome('C:/Users/imran.hossain/Downloads/chromedriver_win32/chromedriver.exe')
# cursor.execute("EXEC dbo.RR_ORMList_Get")
for row in cursor:
    print(row)
#     # webbrowser.get(chromedir).open("https://www.google.com/search?q=" + row[1])
#     # webbrowser.open("https://www.google.com/search?q=" + row[1])
    driver.get("https://www.youtube.com/")
    que = driver.find_element_by_xpath("/html/body/ytd-app/div/div/ytd-masthead/div[3]/div[2]/ytd-searchbox/form/div/div[1]/input")
    que.send_keys(row[1])
    time.sleep(2)
    que.send_keys(Keys.ARROW_DOWN)
    que.send_keys(Keys.ARROW_DOWN)
    time.sleep(2)
    que.send_keys(Keys.RETURN)


# cursor.execute('''
#                 INSERT INTO eOperations.dbo.RR_RiskMaster (RiskTrackingId, RiskTitle, RiskType, RiskTypeOther,
#                 RiskDescription,DateOfRiskDiscovery,DateOfRiskReport,RiskArea,ResponsibleDepartment,ResponsibleDepartmentOther,ResponsiblePerson,
#                 IsRecurring,RecurringRiskDescription,HasOperationalLoss,OperationalLossDescription,TentativeRootCause,TentativeRootCauseOthers,
#                 BusinessImpact,BusinessImpactOthers,IsActionTaken,ActionTakenDescription,ProvidersPin,Recommendation,ProposedClosingDate,
#                 Status,Remarks,ReportedBy,ReportedDate,UpdatedBy,UpdatedDate,UpdatedOn,SubmittedDate,IsDelete,DeletedBy,DeletedDate,
#                 DeletedRemarks,Segment)
#                 VALUES
#                 ('RMU210111Imran','RiskTitlerisk-deba-1','ORM','','RiskDescriptionabcd','2020-1-11',
#                 '2020-1-12','RiskAreaabc','ResponsibleDepartmentabchdh','','',0,'',0,'','AOP','','REM','',0,'',
#                 '1324','Recommendation','2021-03-01','STOO3','','imran.hossain','2021-3-1',
#                 '','','','',0,'','','','RET')
#
#                 ''')
# conn.commit()

conn.commit()


