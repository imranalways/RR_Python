import pyodbc
import ConfigurationManager
import EncryptDecrypt


ServerName = EncryptDecrypt.Decrypt(ConfigurationManager.ConnectionString("Host"))
Database = EncryptDecrypt.Decrypt(ConfigurationManager.ConnectionString("ServiceName"))
UserId = EncryptDecrypt.Decrypt(ConfigurationManager.ConnectionString("UserName"))
Password = EncryptDecrypt.Decrypt(ConfigurationManager.ConnectionString("Password"))
string = 'DRIVER={SQL Server};SERVER=' + ServerName.strip() + ';DATABASE=' + Database.strip() + ';UID=' + UserId.strip() + ';PWD=' + Password.strip() + ';'


def ExecuteStoredProcedure(procedureName):
    conn = pyodbc.connect(string)
    with conn:
        with conn.cursor() as crs:
            crs.execute(procedureName)
            columns = [column[0] for column in crs.description]
            results = []
            for row in crs.fetchall():
                results.append(dict(zip(columns, row)))
            crs.commit()
            del crs
    conn.close()
    return results


def ExecuteNonQueryStoredProcedure(procedureName, param):
    conn = pyodbc.connect(string)
    paramList = ""
    paramValue = []
    with conn:
        with conn.cursor() as crs:
            for key in param:
                paramList += " @" + key + "=?, "
                paramValue.append(param[key])
            paramList=paramList[:-2]
            procedureName = procedureName + paramList
            crs.execute(procedureName, paramValue)
            crs.commit()
            del crs
    conn.close()
