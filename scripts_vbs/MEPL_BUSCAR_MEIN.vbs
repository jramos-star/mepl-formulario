' ============================================================
' MEPL_BUSCAR_MEIN.vbs
' ------------------------------------------------------------
' Uso:
'   cscript //nologo MEPL_BUSCAR_MEIN.vbs ORDEN
'
' Busca en IW39 la MEIN hija de la orden y la imprime:
'   OK | MEIN: 421557629 | Orden: 421480268
'
' Si no encuentra MEIN imprime:
'   ERROR | Orden: 421480268 | sin MEIN
' ============================================================

Dim orden
Dim SapGuiAuto, application, connection, session

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Uso: cscript //nologo MEPL_BUSCAR_MEIN.vbs ORDEN"
    WScript.Quit 1
End If

orden = Trim(WScript.Arguments(0))

If orden = "" Then
    WScript.Echo "ERROR: falta el numero de orden."
    WScript.Quit 1
End If

' Conectar a SAP GUI
If Not IsObject(application) Then
    Set SapGuiAuto  = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
    Set connection = application.Children(0)
End If
If Not IsObject(session) Then
    Set session = connection.Children(0)
End If
If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject application, "on"
End If

session.findById("wnd[0]").maximize

' Abrir IW39
session.findById("wnd[0]/tbar[0]/okcd").text = "iw39"
session.findById("wnd[0]").sendVKey 0

' Ingresar orden y ejecutar
session.findById("wnd[0]/usr/chkDY_MAB").selected = True
session.findById("wnd[0]/usr/chkDY_HIS").selected = True
session.findById("wnd[0]/usr/ctxtAUFNR-LOW").text = orden
session.findById("wnd[0]/usr/ctxtDATUV").text = ""
session.findById("wnd[0]/usr/ctxtDATUB").text = ""
session.findById("wnd[0]/usr/ctxtAUFNR-LOW").setFocus
session.findById("wnd[0]/usr/ctxtAUFNR-LOW").caretPosition = Len(orden)
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press

WScript.Sleep 500

' Verificar error en barra de estado
Dim msgBar
msgBar = Trim(session.findById("wnd[0]/sbar").text)
If InStr(LCase(msgBar), "no existe") > 0 Or _
   InStr(LCase(msgBar), "no se encontr") > 0 Or _
   InStr(LCase(msgBar), "error") > 0 Then
    WScript.Echo "ERROR | Orden: " & orden & " | SAP: " & msgBar
    ' Volver atras 3 veces para quedar en IW39
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    WScript.Quit 1
End If

' Expandir arbol de ordenes
On Error Resume Next
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/btn%#AUTOTEXT004").press
WScript.Sleep 300

' Seleccionar fila 2 del arbol (MEIN hija)
session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").selectItem "          2", "1"
If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    WScript.Echo "ERROR | Orden: " & orden & " | sin MEIN hija en el arbol"
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    WScript.Quit 1
End If
On Error GoTo 0

session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem "          2", "1"
session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").doubleClickItem "          2", "1"
WScript.Sleep 300

' Leer numero de MEIN
Dim numMEIN
numMEIN = ""
On Error Resume Next
numMEIN = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-AUFNR").text)
If Err.Number <> 0 Then
    Err.Clear
    numMEIN = ""
End If
On Error GoTo 0

' Volver atras 3 veces para dejar SAP en IW39 listo para la siguiente orden
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

If numMEIN = "" Or numMEIN = orden Then
    WScript.Echo "ERROR | Orden: " & orden & " | no se pudo leer MEIN"
    WScript.Quit 1
End If

WScript.Echo "OK | MEIN: " & numMEIN & " | Orden: " & orden
WScript.Quit 0
