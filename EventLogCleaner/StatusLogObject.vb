Public Class StatusLogObject
    Public Property TOOL_NAME As String
    Public Property SERVER As String
    Public Property STATUS As String
    Public Property RUN_TIME As Date
    Public Property NEXT_RUN As Date
    Public Property ADD_SUP_VALUE As String
    Public ReadOnly Property CSVDATA As String
        Get
            Return String.Join(";", TOOL_NAME, SERVER, STATUS, RUN_TIME, NEXT_RUN, ADD_SUP_VALUE)
        End Get
    End Property

End Class
