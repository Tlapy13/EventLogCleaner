Public Class ConfigData
    Public Enum Unit
        Day = 1
        Month = 2
        Year = 3
    End Enum

    Public Property LastRunTime As DateTime = DateTime.Now
    'default connection string is DEV
    Public Property ConnectionString As String = "TsdOdRXrNUveiuQrtIT5UjuyeTWDK1ttvkcnCmyWwBXWB9gDWaf08uhzK65u/O7OR4bhq30d50gaOIAI1/d2orP4OgOeZhlVfnvr6Rc/uWV9LeJqfxBhowHJWqAwpq3pgbkviPza85y0BKbsYknF/aCk3prVW57mo2CLY9o71+c="
    Public Property CleanStatusLogsInterval As Integer = 3
    Public Property CleanEventLogsInterval As Integer = 6
    Public Property BackupEnabled As Boolean = False
    Public Property BackupFolderPath As String = "Backup"
    Public Property ClearEventLog As Boolean = False
    Public Property ClearStatusDataLog As Boolean = False
    Public Property Server As String = "App"
    Public Property NextRunUnit As Unit = Unit.Day
    Public Property NextRunInterval As Integer = 1
    Public ReadOnly Property NextRun As DateTime
        Get
            Dim NextRunDate As Date = DateTime.Now

            Select Case NextRunUnit
                Case Unit.Day
                    NextRunDate = NextRunDate.AddDays(NextRunInterval)
                Case Unit.Month
                    NextRunDate = NextRunDate.AddMonths(NextRunInterval)
                Case Unit.Year
                    NextRunDate = NextRunDate.AddYears(NextRunInterval)
            End Select

            Return NextRunDate
        End Get
    End Property

    Public Sub New()
        LastRunTime = My.Settings.LastRunTime
        CleanStatusLogsInterval = My.Settings.CleanStatusLogsInterval
        CleanEventLogsInterval = My.Settings.CleanEventsLogsInterval
        ConnectionString = My.Settings.ConnectionString
        BackupEnabled = My.Settings.BackupEnabled
        BackupFolderPath = My.Settings.BackupFolderPath
        ClearEventLog = My.Settings.ClearEventLog
        ClearStatusDataLog = My.Settings.ClearStatusDataLog
        Server = My.Settings.Server
        NextRunUnit = My.Settings.NextRunUnit
        NextRunInterval = My.Settings.NextRunInterval
        SetNewLastRun()
    End Sub

    Private Sub SetNewLastRun()
        My.Settings.LastRunTime = DateTime.Now
        My.Settings.Save()
        My.Settings.Reload()
    End Sub

End Class
