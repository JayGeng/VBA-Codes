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
session.findById("wnd[0]/tbar[0]/okcd").text = "/NMM02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = "PTB-112129"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "8130"
session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = "019"
session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = "006"
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2731/txtMARA-BRGEW").text = "3.049"
session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2732/txtMLGN-LHMG1").text = "176"
session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LGBKZ").text = "901"

