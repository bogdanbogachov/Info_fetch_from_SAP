Option Explicit

' Folder selection
Private Sub Get_folder_main_Click()
Dim status As Integer
    
    Folder_main = "Null"
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    status = diaFolder.Show
    If status = 0 Then
        Exit Sub
    End If
    Folder_main = diaFolder.SelectedItems(1)

End Sub

' Range selection
Private Sub Get_range_Click()
    
    On Error Resume Next
    Set eff_Range = Application.InputBox("Select the range", "Effectivity range selection", , , , , , 8)

End Sub

' Once the user form is closed by X button (top right corner), all variables are set to Nothing or to ""
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        CI_PN = ""
        prog = ""
        Set eff_Range = Nothing
        Folder_main = ""
    End If
End Sub

' Submit button
Private Sub Submit_button_Click()
    
    UserForm = False ' this flag is used to stop implementation of SAP module in case the user form is closed
    
    CI_PN = CI_input.Value
    If CI_PN = "" Then
        MsgBox ("Enter all values")
        Exit Sub
    End If
    
    If A.Value = True Then
        prog = "A"
    Else
        If B.Value = False Then
            MsgBox ("Enter all values")
            Exit Sub
        Else
            prog = "B"
        End If
    End If
    
    If eff_Range Is Nothing Then
        MsgBox ("Enter all values")
        Exit Sub
    End If
    
    If Folder_main = "Null" Or Folder_main = "" Then
        MsgBox ("Enter all values")
        Exit Sub
    End If
    
    Unload Me
    UserForm = True

End Sub
