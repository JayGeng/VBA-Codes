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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl06o"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btnBUTTON6").press
session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").setFocus
session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").caretPosition = 7
session.findById("wnd[0]").sendVKey 14
session.findById("wnd[0]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]").sendVKey 8
session.findById("wnd[0]").sendVKey 8
