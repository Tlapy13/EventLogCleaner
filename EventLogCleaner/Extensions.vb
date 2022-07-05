Imports System.Data.Common
Imports System.Runtime.CompilerServices
Imports System.Text

Module Extensions

    <Extension()>
    Public Function ToCSV(ByVal Data As List(Of Object), separator As String) As StringBuilder

        Dim SB As New StringBuilder

        For Each item In Data
            SB.AppendLine(item.CSVDATA)
        Next

        Return SB

    End Function

End Module
