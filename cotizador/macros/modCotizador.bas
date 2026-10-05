Option Explicit

' =====================================================================
'  COTIZADOR DE ARRENDAMIENTO - Macros
'  Todas las celdas se localizan por NOMBRES DEFINIDOS (inp_*, esc_*,
'  par_*, def_*, res_*), por lo que puede mover filas sin romper macros.
' =====================================================================

Private Const HOJA_COT As String = "Cotizador"
Private Const HOJA_CARTA As String = "Carta Cotizacion"
Private Const HOJA_TABLA As String = "Tabla de Pagos"
Private Const HOJA_HIST As String = "Historial"
Private Const COL_SNAPSHOT As Long = 22      ' columna V: inicio de la copia de datos capturados
Private Const FILA_ENCABEZADO As Long = 4

' Celdas de captura que se guardan y se restauran
Private Function NombresEntrada() As Variant
    NombresEntrada = Array("inp_Folio", "inp_Fecha", "inp_Cliente", "inp_RFC", "inp_Contacto", "inp_Correo", _
        "inp_Proveedor", "inp_Equipo", "inp_Promotor", "inp_Obligado", "inp_Tipo", "inp_Moneda", "inp_TC", _
        "inp_Modalidad", "inp_Valor", "inp_FechaFirma", "inp_FechaPrimera", "inp_ComBanco", "inp_Fondeo", _
        "inp_CMobj", "esc_Incluir", "esc_Plazo", "esc_Sucesivo", "esc_Tasa", "esc_TasaSuc", "esc_Enganche", _
        "esc_Residual", "esc_PagoFinal", "esc_Comision", "esc_Deposito", "esc_Seguro", "esc_SeguroForma", _
        "esc_GPS", "esc_GPSForma")
End Function

Private Function R(ByVal nombre As String) As Range
    Set R = ThisWorkbook.Names(nombre).RefersToRange
End Function

Private Function Valor(ByVal nombre As String) As Variant
    Valor = R(nombre).Cells(1, 1).Value
End Function

Private Function SiguienteFolio() As String
    SiguienteFolio = CStr(Valor("par_Prefijo")) & Format(CLng(Val(Valor("par_Consecutivo"))) + 1, "0000")
End Function

Private Function Usuario() As String
    On Error Resume Next
    Usuario = Application.UserName
    If Usuario = "" Then Usuario = Environ("USERNAME")
End Function

Private Function BuscarFolio(ByVal ws As Worksheet, ByVal folio As String) As Long
    Dim i As Long, ultima As Long
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = FILA_ENCABEZADO + 1 To ultima
        If StrComp(Trim(CStr(ws.Cells(i, 1).Value)), Trim(folio), vbTextCompare) = 0 Then
            BuscarFolio = i
            Exit Function
        End If
    Next i
    BuscarFolio = 0
End Function

' ---------------------------------------------------------------------
'  NUEVA COTIZACIÓN
' ---------------------------------------------------------------------
Public Sub NuevaCotizacion()
    If MsgBox("¿Iniciar una nueva cotización?" & vbCrLf & _
              "Se borrarán los datos capturados que no haya guardado en el Historial.", _
              vbQuestion + vbYesNo, "Nueva cotización") <> vbYes Then Exit Sub
    On Error GoTo Falla
    Application.ScreenUpdating = False
    LimpiarCaptura
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(HOJA_COT).Activate
    R("inp_Cliente").Select
    Exit Sub
Falla:
    Application.ScreenUpdating = True
    MsgBox "No se pudo iniciar la cotización: " & Err.Description, vbExclamation
End Sub

' Pone la captura en blanco con los valores por defecto de Configuracion
Public Sub LimpiarCaptura(Optional ByVal sinDialogo As Boolean = True)

    R("inp_Folio").Value = SiguienteFolio()
    R("inp_Fecha").Value = Date
    R("inp_Cliente").Value = ""
    R("inp_RFC").Value = ""
    R("inp_Contacto").Value = ""
    R("inp_Correo").Value = ""
    R("inp_Proveedor").Value = ""
    R("inp_Equipo").Value = ""
    R("inp_Obligado").Value = "Por definir"
    R("inp_Tipo").Value = Valor("def_Tipo")
    R("inp_Moneda").Value = Valor("def_Moneda")
    R("inp_TC").Value = 1
    R("inp_Modalidad").Value = Valor("def_Modalidad")
    R("inp_Valor").Value = 0
    R("inp_FechaFirma").Value = Date
    R("inp_FechaPrimera").Value = Date
    R("inp_ComBanco").Value = 0
    R("inp_Fondeo").Value = Valor("par_Fondeo")
    R("inp_CMobj").Value = Valor("par_CMmin")

    R("esc_Incluir").Value = "Sí"
    R("esc_Plazo").Value = R("def_Plazo").Value
    R("esc_Sucesivo").Value = R("def_Sucesivo").Value
    R("esc_Tasa").Value = R("def_Tasa").Value
    R("esc_TasaSuc").Value = R("def_TasaSuc").Value
    R("esc_Enganche").Value = R("def_Enganche").Value
    R("esc_Residual").Value = R("def_Residual").Value
    R("esc_PagoFinal").ClearContents
    R("esc_Comision").Value = R("def_Comision").Value
    R("esc_Deposito").Value = R("def_Deposito").Value
    R("esc_Seguro").Value = 0
    R("esc_SeguroForma").Value = "Financiado"
    R("esc_GPS").Value = 0
    R("esc_GPSForma").Value = "Financiado"
End Sub

' ---------------------------------------------------------------------
'  GUARDAR EN HISTORIAL
' ---------------------------------------------------------------------
Public Sub GuardarCotizacion()
    Dim ws As Worksheet, fila As Long, f As Long, nm As Variant, c As Range
    Dim col As Long, k As Long, plazos As String, esNuevo As Boolean

    If Trim(CStr(Valor("inp_Cliente"))) = "" Then
        MsgBox "Capture el nombre del cliente antes de guardar.", vbExclamation: Exit Sub
    End If
    If Val(Valor("inp_Valor")) <= 0 Then
        MsgBox "Capture el valor del equipo antes de guardar.", vbExclamation: Exit Sub
    End If
    If Trim(CStr(Valor("inp_Folio"))) = "" Then R("inp_Folio").Value = SiguienteFolio()

    Set ws = ThisWorkbook.Worksheets(HOJA_HIST)
    f = BuscarFolio(ws, CStr(Valor("inp_Folio")))
    If f = 0 Then
        fila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If fila <= FILA_ENCABEZADO Then fila = FILA_ENCABEZADO + 1
        esNuevo = True
    Else
        If MsgBox("El folio " & Valor("inp_Folio") & " ya existe en el Historial. ¿Desea actualizarlo?", _
                  vbQuestion + vbYesNo, "Guardar cotización") <> vbYes Then Exit Sub
        fila = CLng(f)
    End If

    Application.ScreenUpdating = False
    plazos = ""
    For k = 1 To 4
        If R("esc_Incluir").Cells(1, k).Value = "Sí" Then
            plazos = plazos & IIf(plazos = "", "", "/") & R("esc_Plazo").Cells(1, k).Value
        End If
    Next k

    ws.Cells(fila, 1).Value = CStr(Valor("inp_Folio"))
    ws.Cells(fila, 2).Value = Valor("inp_Fecha")
    ws.Cells(fila, 3).Value = Valor("inp_Cliente")
    ws.Cells(fila, 4).Value = Valor("inp_Equipo")
    ws.Cells(fila, 5).Value = Valor("inp_Promotor")
    ws.Cells(fila, 6).Value = Valor("inp_Moneda")
    ws.Cells(fila, 7).Value = Valor("inp_Valor")
    ws.Cells(fila, 8).Value = "'" & plazos
    For k = 1 To 4
        If R("esc_Incluir").Cells(1, k).Value = "Sí" Then
            ws.Cells(fila, 8 + k).Value = R("res_RentaIVA").Cells(1, k).Value
            ws.Cells(fila, 12 + k).Value = R("res_CMpct").Cells(1, k).Value
        Else
            ws.Cells(fila, 8 + k).ClearContents
            ws.Cells(fila, 12 + k).ClearContents
        End If
    Next k
    If esNuevo Then ws.Cells(fila, 17).Value = "Enviada"
    ws.Cells(fila, 19).Value = Usuario()
    ws.Cells(fila, 20).Value = Now

    ws.Cells(fila, 2).NumberFormat = "dd/mm/yyyy"
    ws.Cells(fila, 7).NumberFormat = "$#,##0.00"
    ws.Range(ws.Cells(fila, 9), ws.Cells(fila, 12)).NumberFormat = "$#,##0.00"
    ws.Range(ws.Cells(fila, 13), ws.Cells(fila, 16)).NumberFormat = "0.00%"
    ws.Cells(fila, 20).NumberFormat = "dd/mm/yyyy hh:mm"

    ' Copia completa de la captura (para poder recargarla)
    col = COL_SNAPSHOT
    For Each nm In NombresEntrada()
        For Each c In R(CStr(nm)).Cells
            If c.MergeArea.Cells(1, 1).Address = c.Address Then
                ws.Cells(FILA_ENCABEZADO, col).Value = CStr(nm) & "|" & (c.Column - R(CStr(nm)).Column + 1)
                ws.Cells(fila, col).Value = c.Value
                col = col + 1
            End If
        Next c
    Next nm

    ' Avanza el consecutivo si se usó el siguiente folio
    If CStr(Valor("inp_Folio")) = SiguienteFolio() Then
        R("par_Consecutivo").Value = CLng(Val(Valor("par_Consecutivo"))) + 1
    End If
    Application.ScreenUpdating = True
    MsgBox "Cotización " & Valor("inp_Folio") & " guardada en el Historial (fila " & fila & ").", vbInformation
End Sub

' ---------------------------------------------------------------------
'  CARGAR DESDE HISTORIAL
' ---------------------------------------------------------------------
Public Sub CargarCotizacion()
    Dim ws As Worksheet, folio As String, f As Long, fila As Long

    Set ws = ThisWorkbook.Worksheets(HOJA_HIST)
    If ActiveSheet.Name = HOJA_HIST And ActiveCell.Row > FILA_ENCABEZADO Then
        folio = CStr(ws.Cells(ActiveCell.Row, 1).Value)
    End If
    folio = InputBox("Folio a cargar:", "Cargar cotización", folio)
    If Trim(folio) = "" Then Exit Sub
    f = BuscarFolio(ws, folio)
    If f = 0 Then
        MsgBox "No se encontró el folio " & folio & " en el Historial.", vbExclamation: Exit Sub
    End If
    fila = CLng(f)
    If MsgBox("Se reemplazará la captura actual con el folio " & folio & ". ¿Continuar?", _
              vbQuestion + vbYesNo, "Cargar cotización") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    RestaurarFila fila
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(HOJA_COT).Activate
    MsgBox "Cotización " & folio & " cargada.", vbInformation
End Sub

' Copia a la hoja Cotizador los datos guardados en una fila del Historial
Public Sub RestaurarFila(ByVal fila As Long)
    Dim ws As Worksheet, col As Long, enc As String, partes() As String, destino As Range
    Set ws = ThisWorkbook.Worksheets(HOJA_HIST)
    col = COL_SNAPSHOT
    Do While Trim(CStr(ws.Cells(FILA_ENCABEZADO, col).Value)) <> ""
        enc = CStr(ws.Cells(FILA_ENCABEZADO, col).Value)
        partes = Split(enc, "|")
        If UBound(partes) = 1 Then
            On Error Resume Next
            Set destino = Nothing
            Set destino = R(partes(0)).Cells(1, CLng(partes(1)))
            On Error GoTo 0
            If Not destino Is Nothing Then destino.Value = ws.Cells(fila, col).Value
        End If
        col = col + 1
    Loop
End Sub

' ---------------------------------------------------------------------
'  PDF Y CORREO
' ---------------------------------------------------------------------
Private Function LimpiarNombre(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbCr, vbLf)
        s = Replace(s, ch, "")
    Next ch
    s = Trim(s)
    If Len(s) > 60 Then s = Left(s, 60)
    LimpiarNombre = Replace(s, " ", "_")
End Function

Private Function CarpetaSalida() As String
    Dim c As String
    c = Trim(CStr(Valor("par_CarpetaPDF")))
    If c = "" Then c = ThisWorkbook.Path
    If c = "" Then c = CurDir
    If Right(c, 1) = "\" Or Right(c, 1) = "/" Then c = Left(c, Len(c) - 1)
    CarpetaSalida = c
End Function

Private Function ExportarHojaPDF(ByVal hoja As String, ByVal prefijo As String, ByVal abrir As Boolean) As String
    Dim ruta As String, sep As String
    sep = Application.PathSeparator
    ruta = CarpetaSalida() & sep & prefijo & "_" & LimpiarNombre(CStr(Valor("inp_Folio"))) & "_" & _
           LimpiarNombre(CStr(Valor("inp_Cliente"))) & ".pdf"
    ThisWorkbook.Worksheets(hoja).ExportAsFixedFormat Type:=xlTypePDF, Filename:=ruta, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=abrir
    ExportarHojaPDF = ruta
End Function

Private Function ValidarParaEnviar() As Boolean
    Dim k As Long, alguno As Boolean, msg As String
    For k = 1 To 4
        If R("esc_Incluir").Cells(1, k).Value = "Sí" Then
            alguno = True
            If R("res_Dictamen").Cells(1, k).Value <> "CUMPLE" Then
                msg = msg & "  - Escenario " & k & ": cash margin " & _
                      Format(R("res_CMpct").Cells(1, k).Value, "0.00%") & " (objetivo " & _
                      Format(Valor("inp_CMobj"), "0.00%") & ")" & vbCrLf
            End If
        End If
    Next k
    If Not alguno Then
        MsgBox "No hay escenarios marcados con ""Incluir en la carta = Sí"".", vbExclamation
        Exit Function
    End If
    If Trim(CStr(Valor("inp_Cliente"))) = "" Then
        MsgBox "Capture el nombre del cliente.", vbExclamation
        Exit Function
    End If
    If msg <> "" Then
        If MsgBox("Atención: los siguientes escenarios NO cumplen el cash margin objetivo:" & vbCrLf & msg & _
                  vbCrLf & "¿Desea continuar de todos modos?", vbExclamation + vbYesNo, "Cash margin") <> vbYes Then
            Exit Function
        End If
    End If
    ValidarParaEnviar = True
End Function

Public Sub ExportarCartaPDF()
    Dim ruta As String
    If Not ValidarParaEnviar() Then Exit Sub
    On Error GoTo Falla
    ruta = ExportarHojaPDF(HOJA_CARTA, "Cotizacion", True)
    MsgBox "Carta guardada en:" & vbCrLf & ruta, vbInformation
    Exit Sub
Falla:
    MsgBox "No se pudo generar el PDF: " & Err.Description & vbCrLf & _
           "Verifique que el archivo esté guardado en una carpeta con permiso de escritura.", vbExclamation
End Sub

Public Sub ExportarTablaPDF()
    Dim ruta As String, esc As Variant
    esc = InputBox("¿Qué escenario desea en la tabla de pagos? (1 a 4)", "Tabla de pagos", Valor("inp_EscTabla"))
    If esc = "" Then Exit Sub
    If Not IsNumeric(esc) Then Exit Sub
    If CLng(esc) < 1 Or CLng(esc) > 4 Then Exit Sub
    R("inp_EscTabla").Value = CLng(esc)
    On Error GoTo Falla
    ruta = ExportarHojaPDF(HOJA_TABLA, "TablaPagos_Esc" & CLng(esc), True)
    MsgBox "Tabla de pagos guardada en:" & vbCrLf & ruta, vbInformation
    Exit Sub
Falla:
    MsgBox "No se pudo generar el PDF: " & Err.Description, vbExclamation
End Sub

Public Sub EnviarPorCorreo()
    Dim ruta As String, ol As Object, m As Object, cuerpo As String
    If Not ValidarParaEnviar() Then Exit Sub
    On Error GoTo FallaPDF
    ruta = ExportarHojaPDF(HOJA_CARTA, "Cotizacion", False)
    On Error GoTo FallaCorreo
    Set ol = CreateObject("Outlook.Application")
    Set m = ol.CreateItem(0)
    cuerpo = "Estimado(a) " & IIf(Trim(CStr(Valor("inp_Contacto"))) = "", "cliente", Valor("inp_Contacto")) & ":" & vbCrLf & vbCrLf & _
             "Por medio del presente le compartimos la cotización de " & Valor("inp_Tipo") & " para " & _
             Valor("inp_Equipo") & " (folio " & Valor("inp_Folio") & ")." & vbCrLf & vbCrLf & _
             "Quedamos atentos a sus comentarios." & vbCrLf & vbCrLf & _
             "Saludos cordiales," & vbCrLf & Valor("inp_Promotor") & vbCrLf & Valor("par_Arrendador")
    With m
        .To = CStr(Valor("inp_Correo"))
        .CC = CStr(Valor("par_CorreoCC"))
        .Subject = "Cotización de arrendamiento " & Valor("inp_Folio") & " - " & Valor("inp_Cliente")
        .Body = cuerpo
        .Attachments.Add ruta
        .Display
    End With
    Exit Sub
FallaPDF:
    MsgBox "No se pudo generar el PDF: " & Err.Description, vbExclamation
    Exit Sub
FallaCorreo:
    MsgBox "No se pudo abrir Outlook (" & Err.Description & ")." & vbCrLf & _
           "El PDF quedó guardado en:" & vbCrLf & ruta, vbExclamation
End Sub

' ---------------------------------------------------------------------
'  AJUSTES FINANCIEROS
' ---------------------------------------------------------------------
Private Function PedirEscenario(ByVal titulo As String, ByVal permitirTodos As Boolean) As Long
    Dim s As Variant, msg As String
    msg = "Número de escenario (1 a 4)" & IIf(permitirTodos, "; 0 = todos los incluidos", "") & ":"
    s = InputBox(msg, titulo, IIf(permitirTodos, 0, 1))
    PedirEscenario = -1
    If s = "" Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    If CLng(s) < IIf(permitirTodos, 0, 1) Or CLng(s) > 4 Then Exit Function
    PedirEscenario = CLng(s)
End Function

Public Sub AplicarTasaMinima()
    Dim esc As Long, k As Long, t As Variant, resumen As String
    esc = PedirEscenario("Aplicar tasa mínima para el cash margin objetivo", True)
    If esc < 0 Then Exit Sub
    For k = 1 To 4
        If (esc = 0 And R("esc_Incluir").Cells(1, k).Value = "Sí") Or esc = k Then
            t = AplicarTasaMinimaEscenario(k)
            If IsNumeric(t) Then
                resumen = resumen & "Escenario " & k & ": " & Format(t, "0.00%") & vbCrLf
            Else
                resumen = resumen & "Escenario " & k & ": no se pudo calcular" & vbCrLf
            End If
        End If
    Next k
    Application.Calculate
    If resumen = "" Then resumen = "No se modificó ningún escenario."
    MsgBox "Tasa anual aplicada (cash margin objetivo " & Format(Valor("inp_CMobj"), "0.00%") & "):" & _
           vbCrLf & vbCrLf & resumen, vbInformation, "Tasa mínima"
End Sub

' Escribe en el escenario k la tasa mínima (redondeada hacia arriba a 0.01%) y la devuelve
Public Function AplicarTasaMinimaEscenario(ByVal k As Long) As Variant
    Dim t As Variant
    Application.Calculate
    t = R("res_TasaMin").Cells(1, k).Value
    If IsNumeric(t) Then
        t = Application.WorksheetFunction.RoundUp(CDbl(t), 4)
        R("esc_Tasa").Cells(1, k).Value = t
        Application.Calculate
    End If
    AplicarTasaMinimaEscenario = t
End Function

Public Sub TasaParaRentaDeseada()
    Dim esc As Long, renta As Variant, ws As Worksheet, n As Double, M As Double, VR As Double, tipo As Long
    Dim tasa As Double
    esc = PedirEscenario("Tasa para la renta deseada", False)
    If esc < 1 Then Exit Sub
    renta = InputBox("Renta mensual deseada SIN IVA para el escenario " & esc & ":", "Renta deseada")
    If renta = "" Then Exit Sub
    If Not IsNumeric(renta) Then MsgBox "Capture un número.", vbExclamation: Exit Sub
    Set ws = ThisWorkbook.Worksheets("Corrida " & esc)
    n = ws.Range("C5").Value
    M = ws.Range("C19").Value
    VR = ws.Range("C21").Value
    tipo = ws.Range("C7").Value
    On Error GoTo Falla
    tasa = Application.WorksheetFunction.Rate(n, CDbl(renta), -M, VR, tipo) * 12
    On Error GoTo 0
    If tasa < 0 Then
        MsgBox "Con esa renta la tasa resultaría negativa (" & Format(tasa, "0.00%") & "). Revise el monto.", vbExclamation
        Exit Sub
    End If
    If MsgBox("La tasa anual que produce una renta de " & Format(renta, "$#,##0.00") & " es " & _
              Format(tasa, "0.0000%") & "." & vbCrLf & vbCrLf & "¿Aplicarla al escenario " & esc & "?", _
              vbQuestion + vbYesNo, "Tasa para renta deseada") = vbYes Then
        R("esc_Tasa").Cells(1, esc).Value = tasa
    End If
    Exit Sub
Falla:
    MsgBox "No existe una tasa que produzca esa renta con los datos actuales.", vbExclamation
End Sub

Public Sub CopiarEscenario1()
    Dim nm As Variant, k As Long
    If MsgBox("¿Copiar sucesivo, tasas, enganche, residual, pago final, comisión, depósito, seguro y GPS del " & _
              "escenario 1 a los escenarios 2, 3 y 4? (los plazos no cambian)", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    For Each nm In Array("esc_Sucesivo", "esc_Tasa", "esc_TasaSuc", "esc_Enganche", "esc_Residual", "esc_PagoFinal", _
                         "esc_Comision", "esc_Deposito", "esc_Seguro", "esc_SeguroForma", "esc_GPS", "esc_GPSForma")
        For k = 2 To 4
            R(CStr(nm)).Cells(1, k).Value = R(CStr(nm)).Cells(1, 1).Value
        Next k
    Next nm
End Sub

' ---------------------------------------------------------------------
'  PROTECCIÓN (sin contraseña: sólo evita borrados accidentales)
' ---------------------------------------------------------------------
Private Function HojasProtegibles() As Variant
    HojasProtegibles = Array(HOJA_COT, HOJA_CARTA, HOJA_TABLA, "Sensibilidad", _
                             "Corrida 1", "Corrida 2", "Corrida 3", "Corrida 4")
End Function

Public Sub ProtegerHojas()
    Dim h As Variant
    For Each h In HojasProtegibles()
        ThisWorkbook.Worksheets(CStr(h)).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
            AllowFormattingColumns:=True, AllowFormattingRows:=True
    Next h
    MsgBox "Hojas protegidas. Sólo las celdas amarillas se pueden editar.", vbInformation
End Sub

Public Sub DesprotegerHojas()
    Dim h As Variant
    For Each h In HojasProtegibles()
        ThisWorkbook.Worksheets(CStr(h)).Unprotect
    Next h
    MsgBox "Hojas desprotegidas. Recuerde volver a protegerlas.", vbInformation
End Sub

' ---------------------------------------------------------------------
'  NAVEGACIÓN
' ---------------------------------------------------------------------
Public Sub IrCotizador()
    ThisWorkbook.Worksheets(HOJA_COT).Activate
End Sub
Public Sub IrCarta()
    ThisWorkbook.Worksheets(HOJA_CARTA).Activate
End Sub
Public Sub IrTabla()
    ThisWorkbook.Worksheets(HOJA_TABLA).Activate
End Sub
Public Sub IrHistorial()
    ThisWorkbook.Worksheets(HOJA_HIST).Activate
End Sub
Public Sub IrSensibilidad()
    ThisWorkbook.Worksheets("Sensibilidad").Activate
End Sub
Public Sub IrCorrida()
    ThisWorkbook.Worksheets("Corrida 1").Activate
End Sub
Public Sub IrConfiguracion()
    ThisWorkbook.Worksheets("Configuracion").Activate
End Sub
Public Sub IrInicio()
    ThisWorkbook.Worksheets("Inicio").Activate
End Sub
