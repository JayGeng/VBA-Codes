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
session.findById("wnd[0]/tbar[0]/okcd").text = "/NLS24"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRL01S-MATNR").setFocus
session.findById("wnd[0]/usr/ctxtRL01S-MATNR").caretPosition = 11
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = false
session.findById("wnd[2]/usr/chkSCAN_STRING-START").setFocus
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[1,2]").setFocus
session.findById("wnd[3]/usr/lbl[1,2]").caretPosition = 3
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
session.findById("wnd[1]/tbar[0]/btn[13]").press
