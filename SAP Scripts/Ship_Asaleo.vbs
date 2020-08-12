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
session.findById("wnd[0]").sendVKey 6
session.findById("wnd[1]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/tbar[0]/btn[24]").press
session.findById("wnd[2]").sendVKey 8
session.findById("wnd[1]").sendVKey 8
session.findById("wnd[0]").sendVKey 16
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1021/ctxtVTTK-TDLNR").text = "1020960"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1021/ctxtVTTK-TDLNR").caretPosition = 7
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_ID").select
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_ID/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1022/txtVTTK-EXTI2").text = "24"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_ID/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1022/txtVTTK-EXTI2").setFocus
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_ID/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1022/txtVTTK-EXTI2").caretPosition = 2
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_AI").select
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_AI/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1030/ctxtVTTK-ADD01").text = "4024"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_AI/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1030/ctxtVTTK-ADD02").text = "4901"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_AI/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1030/txtVTTK-TEXT1").text = "24"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_AI/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1030/txtVTTK-TEXT3").text = "Martin Brower"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_AI/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1030/txtVTTK-TEXT3").setFocus
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_AI/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1030/txtVTTK-TEXT3").caretPosition = 13
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE").select
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1025/btn*RV56A-ICON_STDIS").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1025/btn*RV56A-ICON_STTEN").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]").sendVKey 11
