Dim orden, textoBreve

orden = WScript.Arguments(0)
textoBreve = WScript.Arguments(1)

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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW36"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/ctxtAUFPAR-PM_AUFART").text = "MEIN"
session.findById("wnd[0]/usr/ctxtCAUFVD-MAUFNR").text = orden
session.findById("wnd[0]/usr/ctxtRC62C-REFNR").text = orden
session.findById("wnd[0]/usr/ctxtRC62C-REFNR").setFocus
session.findById("wnd[0]/usr/ctxtRC62C-REFNR").caretPosition = Len(orden)
session.findById("wnd[0]").sendVKey 0

On Error Resume Next
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
On Error GoTo 0

session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-KTEXT").text = textoBreve

session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/ctxtCAUFVD-INGPR").text = "ECO"

' Leer domicilio antes de guardar
Dim domicilio
On Error Resume Next
domicilio = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subOBJECT:SAPLCOIH:7100/txtRIOT-PLTXT").text
On Error GoTo 0

' Guardar la orden
session.findById("wnd[0]/tbar[0]/btn[11]").press

' Capturar número MEIN de la barra de status
Dim msgStatus, numMEIN
msgStatus = session.findById("wnd[0]/sbar").text

' Extraer solo los dígitos al final del mensaje
Dim partes, i
partes = Split(msgStatus, " ")
For i = UBound(partes) To 0 Step -1
    If IsNumeric(partes(i)) And Len(Trim(partes(i))) > 5 Then
        numMEIN = Trim(partes(i))
        Exit For
    End If
Next

' Ruta fija a DIGITALIZAR
Dim fso, archivo, rutaSalida
rutaSalida = "C:\Users\usuario\Desktop\DIGITALIZAR\mein_creada.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
Set archivo = fso.OpenTextFile(rutaSalida, 8, True)
archivo.WriteLine numMEIN & "|" & orden & "|" & domicilio & "|" & textoBreve
archivo.Close

WScript.Echo "MEIN creada: " & numMEIN & " | Orden: " & orden & " | " & domicilio
