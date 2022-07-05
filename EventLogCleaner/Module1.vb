
Imports System.IO
Imports System.Text

Module Module1
    Public Enum LogLevel
        Info = 1
        War = 2
        Err = 3
    End Enum

    Dim addSupVal As New List(Of String)

    Sub Main()
        Dim arguments As String() = Environment.GetCommandLineArgs()
        Dim config As ConfigData
        Dim silentMode As Boolean = False
        Dim sql As New SQL()
        Dim status As SQL.ToolRunStatus = SQL.ToolRunStatus.Finished

        'we are cheating :-) any argument can be used
        If arguments.Length > 1 Then
            'silent mode
            silentMode = True
        Else
            'logging mode
            silentMode = False
        End If

        WriteMessage(silentMode, "Application started", True)
        WriteMessage(silentMode, "Loading config data")
        config = LoadConfigData(silentMode)

        WriteMessage(silentMode, "Config data loaded")

        'Cleaning event logs
        If config.ClearEventLog Then

            FillAddSupVal("EVENTLOGS CLEANING: ")
            WriteMessage(silentMode, "Starting with event logs cleaning", True)

            If CleanLogs(GetType(EventLogObject),
                         False,
                         config.CleanEventLogsInterval,
                         config) Then
                WriteMessage(silentMode, "Event logs cleaning successfully finished!", True)

            Else
                status = SQL.ToolRunStatus.Failed
                WriteMessage(silentMode, "Event logs cleaning finished with error!", True, LogLevel.Err)
            End If

        Else
            FillAddSupVal("EVENTLOGS CLEANING: Disabled")
            WriteMessage(silentMode, "Cleaning event logs is disabled", True)
        End If

        'Cleaning status logs
        If config.ClearStatusDataLog Then

            WriteMessage(silentMode, "Starting with status logs cleaning", True)
            FillAddSupVal("STATUSLOGS CLEANING: ")

            If CleanLogs(GetType(StatusLogObject),
                         False,
                         config.CleanStatusLogsInterval,
                         config) Then
                WriteMessage(silentMode, "Applications Status logs cleaning successfully finished!", True)

            Else
                status = SQL.ToolRunStatus.Failed
                WriteMessage(silentMode, "Applications Status logs cleaning finished with error!", True, LogLevel.Err)
            End If

        Else
            FillAddSupVal("STATUSLOGS CLEANING: Disabled")
            WriteMessage(silentMode, "Cleaning status logs is disabled", True)

        End If

        'Saving status information to DB
        WriteMessage(silentMode, "Saving tool status into DB")
        If sql.SaveToolStatusInfo(status, config, ConvertAddSupValToText()) Then
            WriteMessage(silentMode, "Status was successfully inserted into DB")
        Else
            status = SQL.ToolRunStatus.Failed
            WriteMessage(silentMode, "Error ocured during saving tool status into DB")
        End If

        WriteMessage(silentMode, "LogCleaner has finished all tasks with status: " & status.ToString())

        If Not silentMode Then
            Console.WriteLine()
            Console.WriteLine("========")
            Console.WriteLine("Finished")
            Console.WriteLine("========")

            'we want to keep app opened
            Console.ReadLine()
        End If

    End Sub

    ''' <summary>
    ''' This function will load data from config file
    ''' </summary>
    ''' <param name="SilentMode"></param>
    ''' <returns></returns>
    Private Function LoadConfigData(SilentMode As Boolean) As ConfigData

        Try
            Return New ConfigData()

        Catch ex As Exception

            WriteMessage(SilentMode, "Configuration was not loaded correctly, application will be terminated!", True, LogLevel.Err)
            Environment.Exit(0)
            'this should not happened as application will be terminated first
            Return Nothing
        End Try

    End Function

    Private Function CleanLogs(ObjType As Type,
                               SilentMode As Boolean,
                               Period As Integer,
                               config As ConfigData) As Boolean

        Dim sql As New SQL
        Dim result As Boolean = False
        Dim CSVSB As New StringBuilder
        Dim ReadyForDeletion As Boolean = False
        Dim lineCount As Integer = 0
        Dim deleted As Boolean = False

        If config.BackupEnabled Then

            WriteMessage(SilentMode, "Selecting data for logs backup")
            WriteMessage(SilentMode, "cleanining interval is: " & Period.ToString & " months", True)
            Try

                CSVSB = sql.SelectDataForBackup(lineCount,
                                                Period,
                                                config.ConnectionString,
                                                ObjType)

            Catch ex As Exception
                FillAddSupVal("backup failed")
                WriteMessage(SilentMode, "data could not be selected: " & ex.ToString, True, LogLevel.Err)
                Return False
            End Try

            If CSVSB.Length > 0 Then
                WriteMessage(SilentMode, String.Format("Found {0} lines, saving data to backup file", lineCount))
                ReadyForDeletion = BackupData(SilentMode, CSVSB, ObjType, config.BackupFolderPath)
            Else
                FillAddSupVal("no data for deletion")
                WriteMessage(SilentMode, "no data for deletion found", True)
                Return True
            End If

        Else
            lineCount = sql.SelectLogsCount(Period, config.ConnectionString, ObjType)
            'backup is currently disabled so we can mark it as done
            WriteMessage(SilentMode, "Backup is currently disabled, therefore will be skiped", True)
            ReadyForDeletion = True

            If lineCount = 0 Then
                FillAddSupVal("no data for deletion")
                WriteMessage(SilentMode, "no data for deletion found", True)
                Return True
            End If

        End If

        If ReadyForDeletion Then
            If config.BackupEnabled Then
                WriteMessage(SilentMode, "Backup was sucessfully done. Data can be deleted now")
            End If

            Try
                deleted = sql.DeleteData(Period,
                                         config.ConnectionString,
                                         lineCount,
                                         ObjType)

            Catch ex As Exception
                deleted = False
                WriteMessage(SilentMode, "Error ocured during data deletion: " & ex.ToString, True, LogLevel.War)
            End Try


            If deleted Then
                result = True
                FillAddSupVal("deleted lines: " & lineCount)
                WriteMessage(SilentMode, "Data were deleted sucessfully. Number of rows deleted: " & lineCount, True)
            Else
                FillAddSupVal("deletion failed")
                WriteMessage(SilentMode, "Data were not deleted, error ocured!", True, LogLevel.War)
                result = False

            End If

        End If

        Return result

    End Function

    Private Function BackupData(silentMode As Boolean,
                                cSVSB As StringBuilder,
                                ObjType As Type,
                                BackupPathRef As String) As Boolean
        Try
            Dim Filename As String = "dummy.csv"
            Dim BackupPath As String = ""

            Select Case ObjType
                Case GetType(StatusLogObject)
                    Filename = "StatusLogs_backup.csv"

                Case GetType(EventLogObject)
                    Filename = "EventLogs_backup.csv"
            End Select

            If Directory.Exists(BackupPathRef) Then
                BackupPath = Path.Combine(BackupPathRef, Filename)
                WriteMessage(silentMode, "Backup file will be stored in path: " & BackupPath, True)
            Else
                BackupPath = Path.Combine(My.Application.Info.DirectoryPath, Filename)
                WriteMessage(silentMode, "Path does not exists, file will be saved in root folder: " & BackupPath, True, LogLevel.War)
            End If

            System.IO.File.WriteAllText(BackupPath, cSVSB.ToString())
            Return True

        Catch ex As Exception
            WriteMessage(silentMode, "CSV backup failed!" & ex.ToString())
            Return False
        End Try

    End Function

    Private Sub WriteMessage(silentMode As Boolean,
                             text As String,
                             Optional logToEventLog As Boolean = False,
                             Optional LogType As LogLevel = LogLevel.Info)
        If Not silentMode Then
            Console.WriteLine(DateTime.Now & " - " & text)
        End If

        If logToEventLog Then

            Select Case LogType
                Case 1
                    logging.LogFs.WriteInformationToLog("LogCleaner tool - " & text)
                Case 2
                    logging.LogFs.WriteWarningToLog("LogCleaner tool - " & text)
                Case Else
                    logging.LogFs.WriteErrorToLog("LogCleaner tool - " & text)
            End Select

        End If

    End Sub

    Private Sub FillAddSupVal(text As String)
        addSupVal.Add(text)
    End Sub

    Private Function ConvertAddSupValToText() As String
        Dim convertedText As String = ""
        convertedText = String.Join(", ", addSupVal)
        Return convertedText.Replace("EVENTLOGS CLEANING: ,", "EVENTLOGS CLEANING: ").Replace("STATUSLOGS CLEANING: , ", "STATUSLOGS CLEANING: ")
    End Function

End Module
