' ============================================================
' MEPL_IW32_RESPUESTAS.vbs
' ------------------------------------------------------------
' Uso:
'   cscript //nologo MEPL_IW32_RESPUESTAS.vbs ORDEN "RESP1||RESP2" "TITULO1||TITULO2" "C:\cert1.pdf||C:\cert2.pdf"
'
' ORDEN            : numero de MEIN (col A del Sheets)
' RESPUESTAS       : bloques separados por || (col E del Sheets)
' TITULOS_CERT     : titulos de los certificados separados por || (col G del Sheets)
'                    opcional — si esta vacio no adjunta nada
' RUTAS_CERT       : rutas completas separadas por || (col H del Sheets)
'                    debe tener la misma cantidad de items que TITULOS_CERT
' ============================================================

Dim orden, respuestasCrudas, titulosCertificados, rutasCertificados
Dim SapGuiAuto, application, connection, session

' ------------------------------------------------------------
' Argumentos
' ------------------------------------------------------------
If WScript.Arguments.Count < 2 Then
    WScript.Echo "Uso: cscript //nologo MEPL_IW32_RESPUESTAS.vbs ORDEN ""RESP1||RESP2"" ""TITULO1||TITULO2"" ""C:\cert1.pdf||C:\cert2.pdf"""
    WScript.Quit 1
End If

orden                = Trim(WScript.Arguments(0))
respuestasCrudas     = Trim(WScript.Arguments(1))
titulosCertificados  = ""
rutasCertificados    = ""
If WScript.Arguments.Count >= 3 Then
    titulosCertificados = Trim(WScript.Arguments(2))
End If
If WScript.Arguments.Count >= 4 Then
    rutasCertificados = Trim(WScript.Arguments(3))
End If

If orden = "" Then
    WScript.Echo "ERROR: falta el numero de orden MEIN."
    WScript.Quit 1
End If

If respuestasCrudas = "" Then
    WScript.Echo "ERROR: falta el campo de respuestas."
    WScript.Quit 1
End If

' ------------------------------------------------------------
' Conectar a SAP GUI
' ------------------------------------------------------------
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

' ============================================================
' IW32 - ABRIR ORDEN MEIN
' ============================================================
session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW32"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = orden
session.findById("wnd[0]").sendVKey 0

' ============================================================
' TAB OPERACIONES - cargar claves y textos
' ============================================================
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpVGUE").select

' Leer claves base y separar respuestas
Dim clavesBase, respuestas
clavesBase = LeerClavesBase()
respuestas = SepararRespuestas(respuestasCrudas)

' Cargar claves y textos en la tabla
' Los popups se cierran dentro de CargarOperaciones al hacer scroll
CargarOperaciones clavesBase, respuestas

' ============================================================
' CAMBIO DE TAB: Operaciones -> Cabecera
' Cerrar cualquier popup remanente
' ============================================================
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpIHKZ").select

Dim k
For k = 1 To 30
    WScript.Sleep 400
    On Error Resume Next
    session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit For
    End If
    On Error GoTo 0
Next

' ============================================================
' TAB CABECERA - fechas y grupo planificador
' ============================================================
Dim fechaHoy
fechaHoy = ObtenerFechaHoy()

' Verificar FENA leyendo la barra antes de escribir fechas
Dim msgBarFENA
On Error Resume Next
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subTERM:SAPLCOIH:7300/ctxtCAUFVD-GSTRP").text = fechaHoy
WScript.Sleep 150
msgBarFENA = Trim(session.findById("wnd[0]/sbar").text)
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0
If InStr(UCase(msgBarFENA), "FENA") > 0 Or InStr(LCase(msgBarFENA), "no se puede") > 0 Then
    session.findById("wnd[0]").sendVKey 12
    WScript.Sleep 150
Else
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subTERM:SAPLCOIH:7300/ctxtCAUFVD-GLTRP").text = fechaHoy
End If
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/ctxtCAUFVD-INGPR").text = "E1D"
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subTERM:SAPLCOIH:7300/ctxtCAUFVD-GLTRP").setFocus
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subTERM:SAPLCOIH:7300/ctxtCAUFVD-GLTRP").caretPosition = 10

' ============================================================
' ESTADOS - EJEC + AINS o SUSP segun respuestas
' ============================================================
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/btn%#AUTOTEXT001").press

If TodasNoAplica(respuestas) Then
    ' SUSP: sin scroll, fila 1
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,1]").selected = True
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,1]").setFocus
Else
    ' EJEC: scroll pos 5 fila 2
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 5
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,2]").selected = True
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,2]").setFocus
    ' AINS: scroll pos 8 fila 1
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 8
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,1]").selected = True
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,1]").setFocus
End If
session.findById("wnd[1]/tbar[0]/btn[0]").press

' ============================================================
' ADJUNTAR CERTIFICADOS PDF (si se recibieron titulos y rutas)
' Se hace ANTES de guardar — SAP lo requiere para poder guardar
' ============================================================
If titulosCertificados <> "" And rutasCertificados <> "" Then
    AdjuntarCertificados titulosCertificados, rutasCertificados
End If

' ============================================================
' VERIFICAR Y GUARDAR
' ============================================================

' btn[25] = Liberar (banderita verde) — si no existe lo saltea
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[25]").press
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0
WScript.Sleep 300
' Cerrar popups que pueda abrir btn[25]
Dim pv
For pv = 1 To 30
    WScript.Sleep 200
    On Error Resume Next
    session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit For
    End If
    On Error GoTo 0
Next

' btn[11] = Grabar
session.findById("wnd[0]/tbar[0]/btn[11]").press

' Verificar mensaje de la barra de estado
Dim msgSAP
WScript.Sleep 500
msgSAP = Trim(session.findById("wnd[0]/sbar").text)

Dim guardadoOK
guardadoOK = InStr(LCase(msgSAP), "grabado") > 0 Or _
             InStr(LCase(msgSAP), "guardado") > 0 Or _
             InStr(LCase(msgSAP), "saved") > 0 Or _
             InStr(LCase(msgSAP), "actualizada") > 0 Or _
             InStr(LCase(msgSAP), "actualizado") > 0 Or _
             msgSAP = ""

If Not guardadoOK Then
    WScript.Echo "ERROR | Orden: " & orden & " | SAP: " & msgSAP
    WScript.Quit 1
End If

' ============================================================
' LIBERAR la orden (banderita verde) si no esta ya liberada
' ============================================================
WScript.Sleep 300
Dim statLiberar
statLiberar = ""
On Error Resume Next
statLiberar = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-STTXT").text)
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0

' Solo liberar si NO tiene ya LIB o AINS en el estado
Dim yaLiberada
yaLiberada = InStr(UCase(statLiberar), "LIB") > 0 Or InStr(UCase(statLiberar), "AINS") > 0

If Not yaLiberada Then
    On Error Resume Next
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/btn%#AUTOTEXT002").press
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    WScript.Sleep 300
    ' Cerrar popups de liberacion
    Dim rl
    For rl = 1 To 10
        WScript.Sleep 200
        On Error Resume Next
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit For
        End If
        On Error GoTo 0
    Next
    ' Guardar despues de liberar
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    WScript.Sleep 300
End If

WScript.Echo "OK | Orden: " & orden
WScript.Quit 0


' ============================================================
' SUB: AdjuntarCertificados
' Recibe titulos y rutas separados por ||
' Los empareja por posicion y llama a AdjuntarUnCertificado
' ============================================================
Sub AdjuntarCertificados(titulosCrudos, rutasCrudas)
    Dim titulos, rutas, i, titulo, ruta
    titulos = Split(titulosCrudos, "||")
    rutas   = Split(rutasCrudas,   "||")

    For i = 0 To UBound(titulos)
        titulo = Trim(titulos(i))
        If i <= UBound(rutas) Then
            ruta = Trim(rutas(i))
        Else
            ruta = ""
        End If

        If titulo <> "" And ruta <> "" Then
            AdjuntarUnCertificado titulo, ruta
        End If
    Next
End Sub


' ============================================================
' SUB: AdjuntarUnCertificado
' Adjunta un archivo PDF a la orden via GOS (Crear Documento Anexo)
' Recibe el titulo descriptivo y la ruta completa del archivo
' ============================================================
Sub AdjuntarUnCertificado(titulo, rutaCompleta)
    Dim fso

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(rutaCompleta) Then
        WScript.Echo "ADVERTENCIA: no se encontro el certificado: " & rutaCompleta
        Exit Sub
    End If

    ' Truncar titulo a 40 caracteres si es necesario
    If Len(titulo) > 40 Then titulo = Left(titulo, 40)

    ' Abrir menu GOS -> Crear Documento Anexo
    session.findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
    session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_CREATE_DOC"

    ' Escribir el titulo y confirmar con btn[0] directamente
    session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").text = titulo
    session.findById("wnd[1]/tbar[0]/btn[0]").press

    ' Escribir la ruta completa en PATH y dejar FILENAME vacio
    ' SAP acepta la ruta completa directamente sin abrir exploradores
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = rutaCompleta
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ""
    session.findById("wnd[1]/tbar[0]/btn[0]").press

    ' Popup de seguridad SAP GUI es una ventana Windows nativa
    ' Activamos la ventana por titulo y presionamos Alt+P (Permitir)
    WScript.Sleep 800
    Dim oShell
    Set oShell = CreateObject("WScript.Shell")
    On Error Resume Next
    oShell.AppActivate "Seguridad SAP GUI"
    WScript.Sleep 300
    oShell.SendKeys "%p"   ' Alt+P = Permitir
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 400

    ' Confirmar tipo de almacenamiento
    session.findById("wnd[1]/usr/btnBUTTON_1").press

    ' Confirmacion final
    session.findById("wnd[1]/tbar[0]/btn[0]").press

    WScript.Sleep 300
End Sub


' ============================================================
' FUNCTION: LeerClavesBase
' Lee las claves que SAP ya tiene en la tabla de operaciones
' ============================================================
Function LeerClavesBase()
    Dim arr(), fila, clave, cantidad
    ReDim arr(0)
    cantidad = 0

    For fila = 0 To 50
        On Error Resume Next
        clave = Trim(session.findById(PathClave(fila)).text)
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        On Error GoTo 0

        If clave = "" Then
            ' nada
        ElseIf clave = "ME-MTTO" Then
            Exit For
        Else
            ReDim Preserve arr(cantidad)
            arr(cantidad) = clave
            cantidad = cantidad + 1
        End If
    Next

    LeerClavesBase = arr
End Function


' ============================================================
' SUB: CargarOperaciones
' Escribe claves y textos en la tabla, mas ME-MTTO al final
' ============================================================
Sub CargarOperaciones(claves, resps)
    Dim cantClaves, cantResps, i, j
    Dim filaAbsoluta, filaRelativa, scrollActual
    Dim resp, lineas
    Dim FILAS_VISIBLES

    FILAS_VISIBLES = 6   ' filas visibles sin scroll
    scrollActual   = 0   ' posicion actual del scrollbar

    cantClaves = UBound(claves) - LBound(claves) + 1
    cantResps  = UBound(resps)  - LBound(resps)  + 1

    ' filaAbsoluta: posicion real en la tabla completa (0-based)
    filaAbsoluta = cantClaves

    For i = 0 To cantResps - 1
        If i > cantClaves - 1 Then Exit For

        resp = resps(i)
        If Trim(resp) = "" Then
            filaAbsoluta = filaAbsoluta + 1
        Else
            lineas = DividirEnLineas(resp, 40)

            ' Verificar si necesitamos scroll antes de escribir
            If filaAbsoluta >= scrollActual + FILAS_VISIBLES Then
                scrollActual = filaAbsoluta - FILAS_VISIBLES + 1
                session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010").verticalScrollbar.position = scrollActual
                ' Cerrar popups que dispara el scroll
                Dim s1
                For s1 = 1 To 15
                    WScript.Sleep 200
                    On Error Resume Next
                    session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
                    If Err.Number <> 0 Then
                        Err.Clear
                        On Error GoTo 0
                        Exit For
                    End If
                    On Error GoTo 0
                Next
            End If

            ' filaRelativa = indice que usa findById segun scroll actual
            filaRelativa = filaAbsoluta - scrollActual

            session.findById(PathClave(filaRelativa)).text = claves(i)
            session.findById(PathTexto(filaRelativa)).text = lineas(0)
            filaAbsoluta = filaAbsoluta + 1

            For j = 1 To UBound(lineas)
                ' Verificar scroll para lineas de continuacion
                If filaAbsoluta >= scrollActual + FILAS_VISIBLES Then
                    scrollActual = filaAbsoluta - FILAS_VISIBLES + 1
                    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010").verticalScrollbar.position = scrollActual
                    Dim s2
                    For s2 = 1 To 15
                        WScript.Sleep 200
                        On Error Resume Next
                        session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
                        If Err.Number <> 0 Then
                            Err.Clear
                            On Error GoTo 0
                            Exit For
                        End If
                        On Error GoTo 0
                    Next
                End If
                filaRelativa = filaAbsoluta - scrollActual
                session.findById(PathTexto(filaRelativa)).text = lineas(j)
                filaAbsoluta = filaAbsoluta + 1
            Next
        End If
    Next

    ' ME-MTTO — verificar scroll antes de escribir
    If filaAbsoluta >= scrollActual + FILAS_VISIBLES Then
        scrollActual = filaAbsoluta - FILAS_VISIBLES + 1
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010").verticalScrollbar.position = scrollActual
        Dim s3
        For s3 = 1 To 15
            WScript.Sleep 200
            On Error Resume Next
            session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Exit For
            End If
            On Error GoTo 0
        Next
    End If

    filaRelativa = filaAbsoluta - scrollActual
    session.findById(PathClave(filaRelativa)).text = "ME-MTTO"
    session.findById(PathDuracion(filaRelativa)).text = "1"
    session.findById(PathDuracion(filaRelativa)).setFocus
End Sub


' ============================================================
' SUB: HacerScrollSiNecesario — NO SE USA
' ============================================================
Sub HacerScrollSiNecesario(filaActual, FILAS_VISIBLES)
End Sub


' ============================================================
' FUNCTION: SepararRespuestas
' ============================================================
Function SepararRespuestas(texto)
    Dim partes, limpias(), i, count
    partes = Split(texto, "||")
    ReDim limpias(0)
    count = 0

    For i = 0 To UBound(partes)
        ReDim Preserve limpias(count)
        limpias(count) = Trim(partes(i))
        count = count + 1
    Next

    SepararRespuestas = limpias
End Function


' ============================================================
' FUNCTION: DividirEnLineas
' ============================================================
Function DividirEnLineas(texto, maxLen)
    Dim palabras, lineaActual, lineas(), count, i, propuesta
    texto    = Trim(texto)
    palabras = Split(texto, " ")
    ReDim lineas(0)
    count       = 0
    lineaActual = ""

    For i = 0 To UBound(palabras)
        If lineaActual = "" Then
            propuesta = palabras(i)
        Else
            propuesta = lineaActual & " " & palabras(i)
        End If

        If Len(propuesta) <= maxLen Then
            lineaActual = propuesta
        Else
            ReDim Preserve lineas(count)
            lineas(count) = lineaActual
            count = count + 1
            lineaActual = palabras(i)
        End If
    Next

    If Trim(lineaActual) <> "" Then
        ReDim Preserve lineas(count)
        lineas(count) = lineaActual
    End If

    DividirEnLineas = lineas
End Function


' ============================================================
' PATHS de la tabla de operaciones
' ============================================================
Function PathClave(fila)
    PathClave = "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/ctxtAFVGD-KTSCH[5," & fila & "]"
End Function

Function PathTexto(fila)
    PathTexto = "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-LTXA1[7," & fila & "]"
End Function

Function PathDuracion(fila)
    PathDuracion = "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-DAUNO[13," & fila & "]"
End Function


' ============================================================
' FUNCTION: TodasNoAplica
' ============================================================
Function TodasNoAplica(resps)
    Dim i
    For i = 0 To UBound(resps)
        If Not EsNoAplica(resps(i)) Then
            TodasNoAplica = False
            Exit Function
        End If
    Next
    TodasNoAplica = True
End Function


' ============================================================
' FUNCTION: EsNoAplica
' ============================================================
Function EsNoAplica(texto)
    Dim t
    t = LCase(Trim(texto))
    EsNoAplica = (t = "no aplica") Or _
                 (t = "n/a")       Or _
                 (t = "no corresponde") Or _
                 (t = "no se realizo") Or _
                 (t = "no se hizo") Or _
                 (t = "") Or _
                 (InStr(t, "obra") > 0)
End Function


' ============================================================
' FUNCTION: ObtenerFechaHoy -> DD.MM.YYYY
' ============================================================
Function ObtenerFechaHoy()
    Dim d, m, a
    d = Day(Date)
    m = Month(Date)
    a = Year(Date)
    If d < 10 Then d = "0" & d
    If m < 10 Then m = "0" & m
    ObtenerFechaHoy = d & "." & m & "." & a
End Function
