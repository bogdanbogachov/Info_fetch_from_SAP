Option Explicit
Sub Find_PNs_in_main()

Dim FindString As String
Dim Rng As Range
Dim rResult As Range
Dim OpenBook As Workbook
Dim diaFolder As FileDialog
Dim selected As Boolean
Dim Folder As String
Dim i As Integer
Dim StrFile As String
Dim firstAddress As String
Dim ws As Worksheet
Dim Check As String
Dim aCell As Range
Dim col As Long
Dim colName As String

' Input box to get a PN to filter
Check = InputBox("Enter a Search value", "PN to filter")
' If condition which allows us to handle the "same sheet names error"
If WorksheetExists(Check & "_In_Main") Then
    MsgBox ("Please specify another value." & vbCrLf & "The value you entered is used to create the next sheet name," & vbCrLf & "while such value was already used before" & vbCrLf & "(excel does not except same sheet names).")
    Exit Sub
Else
    FindString = Check
End If

' If condition which allows us to press cancel on input box
If Trim(FindString) <> "" Then
       
    ' Avoid screen flickering
    Application.ScreenUpdating = False
    
    ' Select a folder where files to be looked at are located
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.Title = "Select a folder where files to be filtered are located"
    selected = diaFolder.Show
    
    ' If condition which allows us to press cancel at a folder selection stage
    If Not selected Then
        Exit Sub
    Else
        ' Add a new sheet to represent a result
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = FindString & "_In_Main"
        
        ' Freeze a top row to always see a column description
        Sheets(ws.Name).Select
        Rows(2).Select
        ActiveWindow.FreezePanes = True
    
        ' i variable allows us to iterate through rows in excel sheet
        i = 3
        
        ' Folder variable stores a path of a selected folder
        Folder = diaFolder.SelectedItems(1)
        
        ' StrFile allows us to iterate through files in a selected folder
        StrFile = Dir(Folder & "\*main*")
        
        ' While loop which iterates through files in a selected folder
        Do While Len(StrFile) > 0
        ' OpenBook variable stores an excel file where we look for our PN
        Set OpenBook = Application.Workbooks.Open(Folder & "\" & StrFile)
        ' Find a column to look at in each file separately
        With OpenBook.Sheets(1)
            Set aCell = .Range("A5:Z5").Find(What:="Part Number", LookIn:=xlValues, lookat:=xlWhole, _
                        MatchCase:=False, SearchFormat:=False)
            If Not aCell Is Nothing Then
                col = aCell.Column
                colName = Split(.Cells(, col).Address, "$")(1)
            Else
                MsgBox ("Column Part Number not found")
            End If
        End With
        With OpenBook.Sheets(1).Range("$" & colName & ":$" & colName)
            ' Rng variable stores a range where a searched PN was found
            Set Rng = .Find(What:="*" & FindString & "*", _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            lookat:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            ' Copy a row with column descriptions
            OpenBook.Sheets(1).Rows(5).Copy
            ' Paste above copied row in this workbook
            ThisWorkbook.Worksheets(ws.Name).Rows(1).PasteSpecial xlPasteValues
            ' Create a thick border to separate column descriptions
            With ThisWorkbook.Worksheets(ws.Name).Rows(1).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = 1
            End With
            ' Insert a #
            ThisWorkbook.Worksheets(ws.Name).Cells(i - 1, 1).Value = "#: " & Left(StrFile, 19)
            ' Loop through the opened workbook and find searched values
            If Not Rng Is Nothing Then
                ' firstAddress is used to stop a loop
                firstAddress = Rng.Address
                Do
                    ' Copy a found value
                    OpenBook.Sheets(1).Rows(Rng.Row).Copy
                    ' Paste the copied value in the current workbook
                    ThisWorkbook.Worksheets(ws.Name).Rows(i).PasteSpecial xlPasteValues
                    If rResult Is Nothing Then
                        Set rResult = Rng
                    Else
                        Set rResult = Union(rResult, Rng)
                    End If
                    ' Find next
                    Set Rng = .FindNext(Rng)
                    i = i + 1
                Loop While Not Rng Is Nothing And Rng.Address <> firstAddress
            End If
        End With
        StrFile = Dir
        ' Set all furhter variable to nothing for a search in the next file
        Set rResult = Nothing
        Set Rng = Nothing
        firstAddress = ""
        ' OpenBook.Application.CutCopyMode = False removes info from clipboard
        OpenBook.Application.CutCopyMode = False
        ' Close the file where a PN was searched
        OpenBook.Close True
        ' Create a thick border to be able to distinguish between #s
        With Worksheets(ws.Name).Rows(i).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 1
        End With
        ' Increase i for the next #
        i = i + 2
        Loop
    End If

    Set diaFolder = Nothing

    ' Avoid screen flickering
    Application.ScreenUpdating = False

End If

End Sub
