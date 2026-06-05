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
session.findById("wnd[0]/tbar[0]/okcd").text = "iw39"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkDY_MAB").selected = true
session.findById("wnd[0]/usr/chkDY_HIS").selected = true
session.findById("wnd[0]/usr/ctxtGEWRK-LOW").text = "ME1-3201"
session.findById("wnd[0]/usr/ctxtDATUV").text = ""
session.findById("wnd[0]/usr/ctxtDATUB").text = ""
session.findById("wnd[0]/usr/ctxtGLTRP-LOW").text = "01.10.2024"
session.findById("wnd[0]/usr/ctxtGLTRP-HIGH").text = "30.06.2026"
session.findById("wnd[0]/usr/ctxtVARIANT").text = "C6JORGE"
session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\usuario\Desktop\FILTRO IW39"
session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 36
session.findById("wnd[1]/tbar[0]/btn[11]").press
