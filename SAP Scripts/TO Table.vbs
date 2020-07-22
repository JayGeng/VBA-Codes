If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/chk[1,3]").selected = true 'Select order Column 1 Row 3?
session.findById("wnd[0]/tbar[1]/btn[8]").press 'Click TO in Foreground
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 18 'Shift-F6
session.findById("wnd[0]/usr/tabsFUNC_TABSTRIP/tabpAQVB/ssubD0106_S:SAPML03T:1061/tblSAPML03TD1061/txtRL03T-SELMG[0,0]").text = "2" 'First item in the table
session.findById("wnd[0]").sendVKey 16 'Shift-F4 Generate Next
session.findById("wnd[0]/usr/tabsFUNC_TABSTRIP/tabpAQVB/ssubD0106_S:SAPML03T:1061/tblSAPML03TD1061/txtRL03T-SELMG[0,1]").text = "1"
session.findById("wnd[0]/usr/tabsFUNC_TABSTRIP/tabpAQVB/ssubD0106_S:SAPML03T:1061/tblSAPML03TD1061/txtRL03T-SELMG[0,2]").text = "1"
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
