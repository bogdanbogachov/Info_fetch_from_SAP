Public eff_Range As Range
Public diaFolder As FileDialog
Public Folder_main As String
Public prog As String
Public CI_PN As String
Public UserForm As Boolean
Sub SAP_fetch()
    
    ' Variable to iterate through effectivities
    Dim i As Range
      
    ' Show user form
    FETCH_SAP.Show
    
    ' Turn off alert messages after each file being saved
    Application.DisplayAlerts = False
    
    ' Exit Sub if user form was closed
    If UserForm = False Then
        Exit Sub
    End If
    
    If Not IsObject(SAPGuiApp) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set SAPGuiApp = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = SAPGuiApp.Children(0)
    End If
    If Not IsObject(SAP_session) Then
       Set SAP_session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject SAP_session, "on"
       WScript.ConnectObject SAPGuiApp, "on"
    End If
    
    For Each i In eff_Range.Cells
        SAP_session.findById("wnd[0]/usr/radGR_NIEO").Select
        SAP_session.findById("wnd[0]/usr/ctxtP_MATNR").text = CI_PN
        SAP_session.findById("wnd[0]/usr/ctxtP_MATNRE").text = prog
        ' var i in the next line is converted from Range to String in
        ' order to avoid a run time error
        SAP_session.findById("wnd[0]/usr/ctxtS_TAILNR").text = CStr(i)
        SAP_session.findById("wnd[0]/tbar[1]/btn[8]").press
        SAP_session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell").pressToolbarContextButton "&MB_EXPORT"
        SAP_session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&PC"
        SAP_session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
        SAP_session.findById("wnd[1]/tbar[0]/btn[0]").press
        SAP_session.findById("wnd[1]/usr/ctxtDY_PATH").text = Folder_main
        SAP_session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = i & "_" & CI_PN & "_main.xls"
        SAP_session.findById("wnd[1]/tbar[0]/btn[0]").press
        SAP_session.findById("wnd[0]/usr/cntlCC_104/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
        SAP_session.findById("wnd[0]/usr/cntlCC_104/shellcont/shell").selectContextMenuItem "&PC"
        SAP_session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
        SAP_session.findById("wnd[1]/tbar[0]/btn[0]").press
        SAP_session.findById("wnd[1]/usr/ctxtDY_PATH").text = Folder_main
        SAP_session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = i & "_" & CI_PN & "_consumption.xls"
        SAP_session.findById("wnd[1]/tbar[0]/btn[0]").press
        SAP_session.findById("wnd[0]/tbar[0]/btn[3]").press
    Next i

End Sub
