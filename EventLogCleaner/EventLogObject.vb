Public Class EventLogObject
    Public Property EVENT_SERVER As String
    Public Property EVENT_LEVEL As String
    Public Property EVENT_TIME As Date
    Public Property EVENT_SOURCE As String
    Public Property EVENT_MESSAGE As String
    Public Property EVENT_ID As Integer
    Public Property CHECKED_BY As String
    Public Property CHECKED_DATE As Date

    Public ReadOnly Property CSVDATA As String
        Get
            Return String.Join(";", EVENT_SERVER, EVENT_LEVEL, EVENT_TIME, EVENT_SOURCE, EVENT_MESSAGE.Replace(vbCr, " / ").Replace(vbLf, " / "), EVENT_ID, CHECKED_BY, CHECKED_DATE)
        End Get
    End Property

End Class
