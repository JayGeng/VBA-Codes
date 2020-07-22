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
session.findById("wnd[0]/usr/ctxtVTTK-TPLST").text = "1730"
session.findById("wnd[0]/usr/cmbVTTK-SHTYP").key = "ZCAR"
session.findById("wnd[0]/usr/ctxtRV56A-SVTRA").text = "cawa"
session.findById("wnd[0]").sendVKey 6
session.findById("wnd[1]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/tbar[0]/btn[24]").press
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1021/ctxtVTTK-TDLNR").text = "10089"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1021/ctxtVTTK-ROUTE").text = "w1w1"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1021/ctxtVTTK-VSBED").text = "5M"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1021/ctxtVTTK-SDABW").text = "SEMI"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1021/ctxtVTTK-EXTI1").text = "cw"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1025/btn*RV56A-ICON_STDIS").press
session.findById("wnd[0]").sendVKey 11

