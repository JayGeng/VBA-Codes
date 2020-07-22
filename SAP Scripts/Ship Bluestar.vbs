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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nvt01n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVTTK-TPLST").text = "1230"
session.findById("wnd[0]/usr/cmbVTTK-SHTYP").key = "ZWBK"
session.findById("wnd[0]/usr/ctxtRV56A-SVTRA").text = "WERKEW"
session.findById("wnd[0]/usr/ctxtRV56A-SVTRA").setFocus
session.findById("wnd[0]/usr/ctxtRV56A-SVTRA").caretPosition = 6
session.findById("wnd[0]").sendVKey 6
session.findById("wnd[1]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/tbar[0]/btn[24]").press
session.findById("wnd[2]").sendVKey 8
session.findById("wnd[1]").sendVKey 8
session.findById("wnd[0]").sendVKey 16
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1021/ctxtVTTK-TDLNR").text = "WBL01"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1021/ctxtVTTK-SDABW").text = "14PL"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1021/ctxtVTTK-SDABW").setFocus
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1021/ctxtVTTK-SDABW").caretPosition = 4
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1025/btn*RV56A-ICON_STDIS").press
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1025/btn*RV56A-ICON_STABF").press
session.findById("wnd[1]/usr/txtMESSTXT1").caretPosition = 38
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
