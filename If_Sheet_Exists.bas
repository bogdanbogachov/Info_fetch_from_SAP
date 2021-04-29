Option Explicit

Function WorksheetExists(ByVal WorksheetName As String) As Boolean
Dim Sht As Worksheet

    For Each Sht In ThisWorkbook.Worksheets
        If Application.Proper(Sht.Name) = Application.Proper(WorksheetName) Then
            WorksheetExists = True
            Exit Function
        End If
    Next Sht
WorksheetExists = False
End Function
