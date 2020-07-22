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
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV50A:2126/ssubCUSTOMER_SCREEN:SAPLZLE_IB_HEAD_TAB:9000/txtWS_LIKP-ZZ_CUST_DATE").text = "08.07.2020"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV50A:2126/ssubCUSTOMER_SCREEN:SAPLZLE_IB_HEAD_TAB:9000/txtWS_LIKP-ZZ_CUST_DATE").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV50A:2126/ssubCUSTOMER_SCREEN:SAPLZLE_IB_HEAD_TAB:9000/txtWS_LIKP-ZZ_CUST_DATE").caretPosition = 10
session.findById("wnd[0]").sendVKey 11
