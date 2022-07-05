Imports System.Data.OleDb
Imports System.IO
Imports System.Text
Imports OracleDatabase


Public Class SQL
    Inherits AbstractDataClass

    Public Enum ToolRunStatus
        Failed
        Finished
    End Enum


    Friend Function SelectDataForBackup(ByRef lineCount As Integer,
                                        period As Integer,
                                        connectionString As String,
                                        objtype As Type) As StringBuilder
        Dim result As New List(Of Object)
        Dim params As New List(Of OleDbParameter)
        Dim SQL As New StringBuilder
        Dim CSVSB As New StringBuilder

        Try

            Using data As New OracleDB(connectionString)
                params.Add(data.MakeInParams("@period", OleDbType.Integer, 4, period))

                Select Case objtype
                    Case GetType(StatusLogObject)
                        SQL = GetSQLforBackupStatusLog()
                        result = GetListResult(SQL, data, params, AddressOf SelectStatusLogDataExtract)

                    Case GetType(EventLogObject)
                        SQL = GetSQLforBackupEventLog()
                        result = GetListResult(SQL, data, params, AddressOf SelectEventLogDataExtract)
                End Select

            End Using

        Catch ex As Exception
            Throw ex
        End Try

        lineCount = result.Count()

        CSVSB = result.ToCSV(";")

        Return CSVSB
    End Function

    Friend Function SelectLogsCount(period As Integer,
                                    connectionString As String,
                                    objtype As Type) As Integer
        Dim result As Integer = 0
        Dim params As New List(Of OleDbParameter)
        Dim dr As OleDbDataReader = Nothing
        Dim SQL As New StringBuilder

        Select Case objtype
            Case GetType(StatusLogObject)
                SQL = GetSQLforSelectCountStatusLog()

            Case GetType(EventLogObject)
                SQL = GetSQLforSelectCountEventLog()
        End Select

        Using data As New OracleDB(connectionString)
            params.Add(data.MakeInParams("@period", OleDbType.Integer, 4, period))

            dr = data.RunSQL_R(SQL.ToString(), params.ToArray())

            Do While dr.Read()
                If Not dr.Item("NumberOfRecords") Is DBNull.Value Then result = dr.Item("NumberOfRecords")
            Loop

        End Using

        Return result

    End Function

    Friend Function DeleteData(period As Integer,
                               connectionString As String,
                               lineCount As Integer,
                               objtype As Type) As Boolean

        Dim params As New List(Of OleDbParameter)
        Dim SQL As New StringBuilder
        Dim count As Integer = 0
        Dim result As Boolean = False

        Select Case objtype
            Case GetType(StatusLogObject)
                SQL = GetSQLforDeleteStatusLog()

            Case GetType(EventLogObject)
                SQL = GetSQLforDeleteEventLog()
        End Select

        Try

            Using Data As New OracleDB(connectionString)

                Data.BeginTransaction()

                params.Add(Data.MakeInParams("@period", OleDbType.Integer, 4, period))

                Try
                    count = Data.RunSQLTrans_W(SQL.ToString, params.ToArray)
                Catch ex As Exception
                    Data.RollBack()
                    Throw ex
                End Try

                If count = lineCount Then
                    Data.Commit()
                    result = True
                Else
                    Data.RollBack()
                    result = False
                    Throw New Exception(String.Format("Deleted lines: {0}, Expected lines: {1}. Deletion failed, rollback was done", count, lineCount))
                End If

            End Using

        Catch ex As Exception
            Throw ex
        End Try

        Return result
    End Function

    Private Function GetSQLforBackupStatusLog() As StringBuilder

        Dim SQL As New StringBuilder
        SQL.AppendLine(" Select * ")
        SQL.AppendLine(" from tool_check_status ")
        SQL.AppendLine(" where ")
        SQL.AppendLine(" trunc(RUN_TIME)  <  trunc(add_months(sysdate,- :period)) ")
        SQL.AppendLine(" and NEXT_RUN < sysdate ")
        SQL.AppendLine(" order by RUN_TIME desc ")

        Return SQL
    End Function

    Private Function GetSQLforBackupEventLog() As StringBuilder

        Dim SQL As New StringBuilder
        SQL.AppendLine(" Select EVENT_SERVER, EVENT_LEVEL, EVENT_TIME, EVENT_SOURCE, dbms_lob.substr(EVENT_MESSAGE) as EVENT_MESSAGE, EVENT_ID, CHECKED_BY, CHECKED_DATE ")
        SQL.AppendLine(" from eventlog ")
        SQL.AppendLine(" where ")
        SQL.AppendLine(" trunc(EVENT_TIME) < trunc(add_months(sysdate,- :period))  ")
        SQL.AppendLine(" order by EVENT_TIME desc ")

        Return SQL
    End Function

    Private Function GetSQLforSelectCountEventLog() As StringBuilder

        Dim SQL As New StringBuilder
        SQL.AppendLine(" Select Count(*) as NumberOfRecords ")
        SQL.AppendLine(" from eventlog ")
        SQL.AppendLine(" where ")
        SQL.AppendLine(" trunc(EVENT_TIME) < trunc(add_months(sysdate,- :period)) ")

        Return SQL
    End Function

    Private Function GetSQLforSelectCountStatusLog() As StringBuilder

        Dim SQL As New StringBuilder
        SQL.AppendLine(" Select Count(*) as NumberOfRecords ")
        SQL.AppendLine(" from tool_check_status ")
        SQL.AppendLine(" where ")
        SQL.AppendLine(" trunc(RUN_TIME)  <  trunc(add_months(sysdate,- :period)) ")
        SQL.AppendLine(" and NEXT_RUN < sysdate ")

        Return SQL
    End Function

    Private Function GetSQLforDeleteStatusLog() As StringBuilder

        Dim SQL As New StringBuilder
        SQL.AppendLine(" Delete from ")
        SQL.AppendLine(" tool_check_status ")
        SQL.AppendLine(" where ")
        SQL.AppendLine(" trunc(RUN_TIME)  <  trunc(add_months(sysdate,- :period)) ")
        SQL.AppendLine(" and NEXT_RUN < sysdate ")

        Return SQL
    End Function

    Private Function GetSQLforDeleteEventLog() As StringBuilder

        Dim SQL As New StringBuilder
        SQL.AppendLine(" Delete from ")
        SQL.AppendLine(" eventlog ")
        SQL.AppendLine(" where ")
        SQL.AppendLine(" trunc(EVENT_TIME) < trunc(add_months(sysdate,- :period)) ")

        Return SQL

    End Function

    Private Function SelectStatusLogDataExtract(dr As OleDbDataReader) As Object
        Dim data As New StatusLogObject

        AssignValueIfPresent(dr, "TOOL_NAME", data.TOOL_NAME)
        AssignValueIfPresent(dr, "SERVER", data.SERVER)
        AssignValueIfPresent(dr, "STATUS", data.STATUS)
        AssignValueIfPresent(dr, "RUN_TIME", data.RUN_TIME)
        AssignValueIfPresent(dr, "NEXT_RUN", data.NEXT_RUN)
        AssignValueIfPresent(dr, "ADD_SUP_VALUE", data.ADD_SUP_VALUE)

        Return data
    End Function

    Private Function SelectEventLogDataExtract(dr As OleDbDataReader) As Object
        Dim data As New EventLogObject

        AssignValueIfPresent(dr, "EVENT_ID", data.EVENT_ID)
        AssignValueIfPresent(dr, "EVENT_LEVEL", data.EVENT_LEVEL)
        AssignValueIfPresent(dr, "EVENT_SOURCE", data.EVENT_SOURCE)
        AssignValueIfPresent(dr, "EVENT_MESSAGE", data.EVENT_MESSAGE)
        AssignValueIfPresent(dr, "EVENT_TIME", data.EVENT_TIME)
        AssignValueIfPresent(dr, "CHECKED_BY", data.CHECKED_BY)
        AssignValueIfPresent(dr, "EVENT_SERVER", data.EVENT_SERVER)
        AssignValueIfPresent(dr, "CHECKED_DATE", data.CHECKED_DATE)

        Return data
    End Function

    Public Function SaveToolStatusInfo(status As ToolRunStatus,
                                  config As ConfigData,
                                  addSupVal As String) As Boolean

        Dim SQL As New StringBuilder
        Dim params As New List(Of OleDbParameter)

        SQL.AppendLine(" INSERT INTO TOOL_CHECK_STATUS ")
        SQL.AppendLine(" (TOOL_NAME, SERVER, STATUS, RUN_TIME, NEXT_RUN, ADD_SUP_VALUE ) ")
        SQL.AppendLine(" VALUES ")
        SQL.AppendLine(" ( ")
        SQL.AppendLine("  :ToolName, ")
        SQL.AppendLine("  :Server, ")
        SQL.AppendLine("  :Status, ")
        SQL.AppendLine("  :RunTime, ")
        SQL.AppendLine("  :NextRun, ")
        SQL.AppendLine("  :Add_Sup_Value ")
        SQL.AppendLine(" ) ")

        Using Data As New OracleDB(config.ConnectionString)

            params.Add(Data.MakeInParams("@ToolName", OleDbType.VarChar, 50, "LogCleaner"))
            params.Add(Data.MakeInParams("@Server", OleDbType.VarChar, 3, config.Server))
            params.Add(Data.MakeInParams("@Status", OleDbType.VarChar, 15, status.ToString))
            params.Add(Data.MakeInParams("@RunTime", OleDbType.Date, 8, DateTime.Now.ToString))
            params.Add(Data.MakeInParams("@NextRun", OleDbType.Date, 8, config.NextRun))
            params.Add(Data.MakeInParams("@Add_Sup_Value", OleDbType.VarChar, 300, If(String.IsNullOrEmpty(addSupVal), "N/A", addSupVal)))

            Try
                Data.RunSQL_W(SQL.ToString, params.ToArray)
            Catch ex As Exception
                Return False
            End Try

        End Using

        Return True

    End Function

End Class
