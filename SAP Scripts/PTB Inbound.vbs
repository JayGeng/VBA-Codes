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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl32n"
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB005").select
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB005/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = "POR1332065"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV50A:2126/ssubCUSTOMER_SCREEN:SAPLZLE_IB_HEAD_TAB:9000/txtWS_LIKP-ZZ_CUST_DATE").text = "07.07.2020"
session.findById("wnd[0]").sendVKey 11
session.findById("wnd[0]/mbar/menu[0]/menu[6]").select
session.findById("wnd[1]/tbar[0]/btn[86]").press
