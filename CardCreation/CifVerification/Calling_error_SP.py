import datetime
import pyodbc
# import yaml
import ConfigurationManager
import DatabaseManager


def error_sp(loanId):
    params = {}
    params["loanId"] = loanId
    DatabaseManager.ExecuteNonQueryStoredProcedure("RDMS_UpdateCIFNoError_Card_Rpa", params)
