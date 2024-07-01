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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nsm37"
session.findById("wnd[0]").sendVKey(0) 
session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "ZU*"
session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "CORE_BASIS2"
session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = "28.06.2024"
session.findById("wnd[0]/usr/ctxtBTCH2170-TO_DATE").text = "28.06.2024"
session.findById("wnd[0]/usr/ctxtBTCH2170-TO_DATE").setFocus
session.findById("wnd[0]/usr/ctxtBTCH2170-TO_DATE").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/lbl[37,13]").setFocus
session.findById("wnd[0]/usr/lbl[37,13]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/usr/lbl[5,3]").setFocus
session.findById("wnd[0]/usr/lbl[5,3]").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "A_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,4]").setFocus
session.findById("wnd[0]/usr/lbl[5,4]").caretPosition = 11
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "C_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,5]").setFocus
session.findById("wnd[0]/usr/lbl[5,5]").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "E_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,6]").setFocus
session.findById("wnd[0]/usr/lbl[5,6]").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "F_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,7]").setFocus
session.findById("wnd[0]/usr/lbl[5,7]").caretPosition = 13
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "G_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,8]").setFocus
session.findById("wnd[0]/usr/lbl[5,8]").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "H_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,9]").setFocus
session.findById("wnd[0]/usr/lbl[5,9]").caretPosition = 14
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "I_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,10]").setFocus
session.findById("wnd[0]/usr/lbl[5,10]").caretPosition = 15
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "J_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,11]").setFocus
session.findById("wnd[0]/usr/lbl[5,11]").caretPosition = 14
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "K_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,12]").setFocus
session.findById("wnd[0]/usr/lbl[5,12]").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "L_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,13]").setFocus
session.findById("wnd[0]/usr/lbl[5,13]").caretPosition = 13
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "M_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,14]").setFocus
session.findById("wnd[0]/usr/lbl[5,14]").caretPosition = 11
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "N_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,15]").setFocus
session.findById("wnd[0]/usr/lbl[5,15]").caretPosition = 13
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "P_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,16]").setFocus
session.findById("wnd[0]/usr/lbl[5,16]").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "S_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,17]").setFocus
session.findById("wnd[0]/usr/lbl[5,17]").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "T_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,18]").setFocus
session.findById("wnd[0]/usr/lbl[5,18]").caretPosition = 14
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "U_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,19]").setFocus
session.findById("wnd[0]/usr/lbl[5,19]").caretPosition = 16
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "V_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,20]").setFocus
session.findById("wnd[0]/usr/lbl[5,20]").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "W_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,21]").setFocus
session.findById("wnd[0]/usr/lbl[5,21]").caretPosition = 11
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "X_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,22]").setFocus
session.findById("wnd[0]/usr/lbl[5,22]").caretPosition = 11
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Z_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,23]").setFocus
session.findById("wnd[0]/usr/lbl[5,23]").caretPosition = 13
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "R_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,24]").setFocus
session.findById("wnd[0]/usr/lbl[5,24]").caretPosition = 14
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "D1_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,25]").setFocus
session.findById("wnd[0]/usr/lbl[5,25]").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "D2_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,26]").setFocus
session.findById("wnd[0]/usr/lbl[5,26]").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "D3_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,27]").setFocus
session.findById("wnd[0]/usr/lbl[5,27]").caretPosition = 13
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "B1_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,28]").setFocus
session.findById("wnd[0]/usr/lbl[5,28]").caretPosition = 11
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "B2_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[5,29]").setFocus
session.findById("wnd[0]/usr/lbl[5,29]").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[0]/usr/lbl[14,3]").setFocus
session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
session.findById("wnd[0]").sendVKey(2)
session.findById("wnd[0]/tbar[1]/btn[48]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "B3_27.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
session.findById("wnd[0]").sendVKey(0)   