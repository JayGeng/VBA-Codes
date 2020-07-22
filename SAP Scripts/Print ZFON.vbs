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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl03n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text = "298362534"
session.findById("wnd[0]/usr/ctxtLIKP-VBELN").caretPosition = 9
session.findById("wnd[0]/mbar/menu[0]/menu[6]").select
session.findById("wnd[1]/usr/tblSAPLVMSGTABCONTROL").getAbsoluteRow(0).selected = true
session.findById("wnd[1]/tbar[0]/btn[86]").press
