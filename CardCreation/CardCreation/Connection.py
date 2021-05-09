import pyodbc


class databaseConnection(object):

    def callProcedure(self, procedureName, connectionString):

        conn = pyodbc.connect(connectionString)
        sqlExecSP = procedureName#"{call RDMS_GetCreditCardInfoBySerialNo_Card_Rpa_py}"
        cursor = conn.cursor()
        cursor.execute(sqlExecSP)

        firstRow = cursor.fetchall()

        columns = [column[0] for column in cursor.description]
        cursor.commit()
        cursor.close()
        results = []
        for row in firstRow:
            results.append(dict(zip(columns, row)))
        return results

    def callUpdateProcedure(self, procedureName, connectionString, loanId, applicationNo):
        conn = pyodbc.connect(connectionString)

        sqlExecSP = procedureName

        cursor = conn.cursor()
        cursor.execute(sqlExecSP, loanId, applicationNo)
        cursor.commit()
        cursor.close()

    def callProcedureSupply(self, procedureName, connectionString, loanId):

        conn = pyodbc.connect(connectionString)
        sqlExecSP = procedureName
        cursor = conn.cursor()
        cursor.execute(sqlExecSP, loanId)

        firstRow = cursor.fetchall()

        columns = [column[0] for column in cursor.description]
        cursor.commit()
        cursor.close()

        results = []
        for row in firstRow:
            results.append(dict(zip(columns, row)))
            # print(results)
        # print(results)
        return results

    def callProcedureError(self, procedureName, connectionString, loanId):

        conn = pyodbc.connect(connectionString)
        sqlExecSP = procedureName
        cursor = conn.cursor()
        cursor.execute(sqlExecSP, loanId)

        cursor.commit()
        cursor.close()

#NOTE: not implemanted in the code
# class fileConnection(object):
#     text_file = open('C:/cardpro_driver/cdriver/Connection.txt', 'r')
#
#     ServerName = text_file.readline()
#     Database = text_file.readline()
#     UserId = text_file.readline()
#     Password = text_file.readline()
#     baseURL = text_file.readlines()
#
#     text_file.close()