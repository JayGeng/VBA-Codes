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
session.findById("wnd[0]/tbar[0]/okcd").text = "/n/scwm/tu"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TU:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_TU:2001/ctxt/SCWM/S_ASPQ_TU-TU_NUM_EXT").setFocus
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TU:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_TU:2001/ctxt/SCWM/S_ASPQ_TU-TU_NUM_EXT").caretPosition = 0
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TU:2000/tabsTABSTRIP_OIP/tabpOK_OIP_TAB1/ssubSUB_OIP_TAB1:/SCWM/SAPLUI_TU:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").pressButton "OK_CREA"
session.findById("wnd[1]/usr/ctxt/SCWM/S_INS_TU-TU_NUM_EXT").text = "AS 134076"
session.findById("wnd[1]/usr/ctxt/SCWM/S_INS_TU-MTR").text = "0001"
session.findById("wnd[1]/usr/ctxt/SCWM/S_INS_TU-PMAT").text = "mtr"
session.findById("wnd[1]/usr/cmb/SCWM/S_INS_TU-ACT_DIR").key = "2"
session.findById("wnd[1]/usr/cmb/SCWM/S_INS_TU-ACT_DIR").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TU:2000/tabsTABSTRIP_OIP/tabpOK_OIP_TAB1/ssubSUB_OIP_TAB1:/SCWM/SAPLUI_TU:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").pressButton "OK_TU_OIP_SEARCH_AND_ASGN_WHR"
session.findById("wnd[1]/usr/txtP1_DOCNO-LOW").text = "134076"
session.findById("wnd[1]/usr/txtP1_DOCNO-LOW").setFocus
session.findById("wnd[1]/usr/txtP1_DOCNO-LOW").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_TU:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB10").select
session.findById("wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_TU:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB10/ssubSUB_ODP1_TAB10:/SCWM/SAPLUI_TU:3300/cntlCONTAINER_TB_ODP1_10/shellcont/shell").pressButton "OK_ASSIGN"
session.findById("wnd[1]/usr/ctxt/SCWM/S_INS_TU_DOOR-DOOR").text = "dr01"
session.findById("wnd[1]/usr/ctxt/SCWM/S_INS_TU_DOOR-DOOR").setFocus
session.findById("wnd[1]/usr/ctxt/SCWM/S_INS_TU_DOOR-DOOR").caretPosition = 4
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TU:2000/tabsTABSTRIP_OIP/tabpOK_OIP_TAB1/ssubSUB_OIP_TAB1:/SCWM/SAPLUI_TU:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").pressButton "OK_ACTIVATE"
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TU:2000/tabsTABSTRIP_OIP/tabpOK_OIP_TAB1/ssubSUB_OIP_TAB1:/SCWM/SAPLUI_TU:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").pressButton "OK_CKECK"
session.findById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TU:2000/tabsTABSTRIP_OIP/tabpOK_OIP_TAB1/ssubSUB_OIP_TAB1:/SCWM/SAPLUI_TU:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").pressButton "SAVE"
