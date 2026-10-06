Option Explicit

' =====================================================================
'  COTIZADOR MASTER FASTPLUS - Macros
'  Todas las celdas se localizan por NOMBRES DEFINIDOS (inp_*, esc_*,
'  par_*, rc_*, res_*), por lo que puede mover filas sin romper macros.
' =====================================================================

Private Const HOJA_COT As String = "Cotizador"
Private Const HOJA_FAC As String = "Factores"
Private Const HOJA_BON As String = "Bonos"
Private Const HOJA_TABLA As String = "Tabla de Pagos"
Private Const HOJA_HIST As String = "Historial"
Private Const HOJA_PROP As String = "Propuesta"
Private Const HOJA_PINI As String = "Pago Inicial"
Private Const HOJA_VENTA As String = "Venta"
Private Const HOJA_COMITE As String = "Riesgo"
Private Const CLAVE_ADMIN As String = "FP-Admin2026"        ' protege libro y hojas; abre Catálogos
Private Const CLAVE_GERENCIA As String = "FP-Gerencia2026"     ' abre Factores, Bonos y Riesgo
Private Const COL_SNAPSHOT As Long = 22      ' columna V: inicio de la copia de datos capturados
Private Const FILA_ENCABEZADO As Long = 4

' Celdas de captura que se guardan y se restauran
Private Function NombresEntrada() As Variant
    NombresEntrada = Array("inp_Folio", "inp_Fecha", "inp_Cliente", "inp_RFC", "inp_Contacto", "inp_Correo", _
        "inp_Proveedor", "inp_Equipo", "inp_Promotor", "inp_Producto", "inp_Obligado", "inp_Tipo", "inp_Moneda", _
        "inp_TC", "inp_Modalidad", "inp_TipoActivo", "inp_Estado", "inp_Originador", "inp_Tier", "inp_Precio", _
        "inp_Descuento", "inp_DiasRP", "inp_IVAEquipo", "inp_UsoAnticipo", "inp_Anticipo", "inp_SeguroFin", _
        "inp_TIIE", "inp_PlazoSol", "inp_Fuente", "inp_FondeoManual", "inp_OtrosMonto", "inp_OtrosDesc", "inp_GPSMonto", _
        "inp_GastosInv", "inp_ComBanco", "inp_CMobj", "inp_ComProm", "inp_SeguroMonto", "inp_Aseguradora", _
        "inp_AseguradoraPor", "esc_Plazo", "esc_Margen", "esc_Residual", "esc_Comision", "esc_DepPct", _
        "esc_Deposito", "esc_Incluir")
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
'  SEGURIDAD: hojas restringidas y contraseñas
' ---------------------------------------------------------------------
Private Function HojasGerencia() As Variant
    HojasGerencia = Array(HOJA_FAC, HOJA_BON, HOJA_COMITE)
End Function

Private Function PedirClave(ByVal nivel As String) As Boolean
    Dim c As Variant
    c = InputBox("Acceso restringido (" & nivel & ")." & vbCrLf & vbCrLf & "Escriba la contraseña:", "FASTPLUS – " & nivel)
    If c = "" Then Exit Function
    If CStr(c) = CLAVE_ADMIN Or (nivel = "Gerencia" And CStr(c) = CLAVE_GERENCIA) Then
        PedirClave = True
    Else
        MsgBox "Contraseña incorrecta.", vbExclamation, "FASTPLUS"
    End If
End Function

Private Sub Visibilidad(ByVal hojas As Variant, ByVal mostrar As Boolean)
    Dim h As Variant
    On Error Resume Next
    ThisWorkbook.Unprotect CLAVE_ADMIN
    For Each h In hojas
        ThisWorkbook.Worksheets(CStr(h)).Visible = IIf(mostrar, xlSheetVisible, xlSheetHidden)
    Next h
    ThisWorkbook.Protect CLAVE_ADMIN, True
End Sub

' Oculta todo lo restringido (se llama al abrir el libro)
Public Sub OcultarRestringidas()
    Dim k As Long
    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets(HOJA_COT).Activate
    Visibilidad Array(HOJA_FAC, HOJA_BON, HOJA_COMITE, "Catalogos", "Corrida 1", "Corrida 2", "Corrida 3", "Corrida 4"), False
    Application.ScreenUpdating = True
End Sub

Public Sub AccesoGerencia()
    If Not PedirClave("Gerencia") Then Exit Sub
    Visibilidad HojasGerencia(), True
    ThisWorkbook.Worksheets(HOJA_FAC).Activate
End Sub

Public Sub OcultarGerencia()
    ThisWorkbook.Worksheets(HOJA_COT).Activate
    Visibilidad HojasGerencia(), False
End Sub

Public Sub AccesoAdministrador()
    If Not PedirClave("Administrador") Then Exit Sub
    Visibilidad Array(HOJA_FAC, HOJA_BON, HOJA_COMITE, "Catalogos", "Corrida 1", "Corrida 2", "Corrida 3", "Corrida 4"), True
    ThisWorkbook.Worksheets("Catalogos").Activate
End Sub

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
    ' Los factores por plazo (hoja Factores) son política comercial y no se borran.
    R("inp_Folio").Value = SiguienteFolio()
    R("inp_Fecha").Value = Date
    R("inp_Cliente").Value = ""
    R("inp_RFC").Value = ""
    R("inp_Contacto").Value = ""
    R("inp_Correo").Value = ""
    R("inp_Proveedor").Value = ""
    R("inp_Equipo").Value = ""
    R("inp_Obligado").Value = "Por definir"
    R("inp_Producto").Value = R("lst_Producto").Cells(1).Value
    R("inp_Tipo").Value = R("lst_Tipo").Cells(1).Value
    R("inp_Moneda").Value = R("lst_Moneda").Cells(1).Value
    R("inp_TC").Value = 1
    R("inp_Modalidad").Value = R("lst_Modalidad").Cells(1).Value
    R("inp_TipoActivo").Value = R("lst_TipoActivo").Cells(1).Value
    R("inp_Estado").Value = "Nuevo"
    R("inp_Originador").Value = R("lst_Originador").Cells(1).Value
    R("inp_Tier").Value = R("rc_Tiers").Cells(R("rc_Tiers").Cells.Count).Value
    R("inp_Precio").Value = 0
    R("inp_Descuento").Value = 0
    R("inp_DiasRP").Value = 0
    R("inp_IVAEquipo").Value = Valor("par_IVA")
    R("inp_UsoAnticipo").Value = R("lst_UsoAnticipo").Cells(1).Value
    R("inp_Anticipo").Value = 0
    R("inp_SeguroFin").Value = "No financiado"
    R("inp_SeguroMonto").Value = 0
    R("inp_Aseguradora").Value = ""
    R("inp_AseguradoraPor").Value = "Cliente"
    R("inp_TIIE").Value = Valor("par_TIIE")
    R("inp_PlazoSol").Value = R("esc_Plazo").Cells(2).Value
    R("inp_Fuente").Value = Valor("par_FuenteDef")
    R("inp_FondeoManual").ClearContents
    R("inp_OtrosMonto").Value = 0
    R("inp_OtrosDesc").Value = ""
    R("inp_GPSMonto").Value = 0
    R("inp_GastosInv").Value = Valor("par_GastosInv")
    R("inp_ComBanco").Value = 0
    R("inp_CMobj").ClearContents
    R("inp_ComProm").Value = 0
End Sub

' ---------------------------------------------------------------------
'  GUARDAR EN HISTORIAL
' ---------------------------------------------------------------------
Public Sub GuardarCotizacion()
    Dim ws As Worksheet, fila As Long, f As Long, nm As Variant, c As Range
    Dim col As Long, k As Long, plazos As String, esNuevo As Boolean, idx As Long

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
        If R("esc_Incluir").Cells(k).Value = "Sí" Then
            plazos = plazos & IIf(plazos = "", "", "/") & R("esc_Plazo").Cells(k).Value
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
        If R("esc_Incluir").Cells(k).Value = "Sí" Then
            ws.Cells(fila, 8 + k).Value = R("res_RentaIVA").Cells(k).Value
            ws.Cells(fila, 12 + k).Value = R("res_CMpct").Cells(k).Value
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
        idx = 0
        For Each c In R(CStr(nm)).Cells
            idx = idx + 1
            ws.Cells(FILA_ENCABEZADO, col).Value = CStr(nm) & "|" & idx
            ws.Cells(fila, col).Value = c.Value
            col = col + 1
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
            Set destino = R(partes(0)).Cells(CLng(partes(1)))
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
    s = Replace(Replace(s, ".", ""), ",", "")
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

' Exporta varias hojas a un solo PDF (en el orden indicado)
Private Function ExportarHojasPDF(ByVal hojas As Variant, ByVal prefijo As String, ByVal abrir As Boolean) As String
    Dim ruta As String, sep As String, actual As Worksheet
    sep = Application.PathSeparator
    ruta = CarpetaSalida() & sep & prefijo & "_" & LimpiarNombre(CStr(Valor("inp_Folio"))) & "_" & _
           LimpiarNombre(CStr(Valor("inp_Cliente"))) & ".pdf"
    Set actual = ActiveSheet
    ThisWorkbook.Activate
    ThisWorkbook.Worksheets(hojas).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ruta, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=abrir
    actual.Select
    ExportarHojasPDF = ruta
End Function

Private Function EscenarioDePlazo(ByVal plazo As Variant) As Long
    Dim k As Long
    EscenarioDePlazo = 0
    If Not IsNumeric(plazo) Then Exit Function
    For k = 1 To R("esc_Plazo").Cells.Count
        If CLng(R("esc_Plazo").Cells(k).Value) = CLng(plazo) Then
            EscenarioDePlazo = k
            Exit Function
        End If
    Next k
End Function

Private Function ListaPlazos() As String
    Dim k As Long
    For k = 1 To R("esc_Plazo").Cells.Count
        ListaPlazos = ListaPlazos & IIf(k = 1, "", ", ") & R("esc_Plazo").Cells(k).Value
    Next k
End Function

' Pregunta el plazo de los documentos y lo deja como "Plazo solicitado"
Private Function ElegirEscenarioDocumentos() As Boolean
    Dim p As Variant
    p = InputBox("Plazo para la propuesta y documentos (" & ListaPlazos() & " meses):", "Plazo solicitado", Valor("inp_PlazoSol"))
    If p = "" Then Exit Function
    If EscenarioDePlazo(p) = 0 Then
        MsgBox "Ese plazo no está en la hoja Factores.", vbExclamation
        Exit Function
    End If
    R("inp_PlazoSol").Value = CLng(p)
    Application.Calculate
    ElegirEscenarioDocumentos = True
End Function

Private Function ValidarParaEnviar() As Boolean
    Dim k As Long
    k = CLng(Valor("inp_EscTabla"))
    If Trim(CStr(Valor("inp_Cliente"))) = "" Then
        MsgBox "Capture el nombre del cliente.", vbExclamation
        Exit Function
    End If
    If R("res_Dictamen").Cells(k).Value <> "CUMPLE" Then
        If MsgBox("Atención: el plazo de " & Valor("inp_PlazoSol") & " meses NO cumple el rate card (" & _
                  R("res_Dictamen").Cells(k).Value & ", margen neto " & _
                  Format(R("res_CMnetoPct").Cells(k).Value, "0.00%") & ")." & vbCrLf & _
                  "¿Desea continuar de todos modos?", vbExclamation + vbYesNo, "Rate card") <> vbYes Then Exit Function
    End If
    ValidarParaEnviar = True
End Function

Public Sub ExportarTablaPDF()
    Dim ruta As String
    If Not ElegirEscenarioDocumentos() Then Exit Sub
    On Error GoTo Falla
    ruta = ExportarHojaPDF(HOJA_TABLA, "TablaPagos_Esc" & Valor("inp_EscTabla"), True)
    MsgBox "Tabla de pagos guardada en:" & vbCrLf & ruta, vbInformation
    Exit Sub
Falla:
    MsgBox "No se pudo generar el PDF: " & Err.Description, vbExclamation
End Sub

Public Sub ExportarPropuestaPDF()
    Dim ruta As String
    If Not ElegirEscenarioDocumentos() Then Exit Sub
    If Not ValidarParaEnviar() Then Exit Sub
    On Error GoTo Falla
    ruta = ExportarHojaPDF(HOJA_PROP, "Propuesta", True)
    MsgBox "Propuesta guardada en:" & vbCrLf & ruta, vbInformation
    Exit Sub
Falla:
    MsgBox "No se pudo generar el PDF: " & Err.Description, vbExclamation
End Sub

' Propuesta + pago inicial + tabla de pagos en un solo PDF
Public Sub ExportarPaqueteCliente()
    Dim ruta As String
    If Not ElegirEscenarioDocumentos() Then Exit Sub
    If Not ValidarParaEnviar() Then Exit Sub
    On Error GoTo Falla
    ruta = ExportarHojasPDF(Array(HOJA_PROP, HOJA_PINI, HOJA_TABLA), "Cotizacion", True)
    MsgBox "Paquete para el cliente guardado en:" & vbCrLf & ruta, vbInformation
    Exit Sub
Falla:
    MsgBox "No se pudo generar el PDF: " & Err.Description, vbExclamation
End Sub

Public Sub ExportarPromesaVentaPDF()
    Dim ruta As String
    If Not ElegirEscenarioDocumentos() Then Exit Sub
    On Error GoTo Falla
    ruta = ExportarHojaPDF(HOJA_VENTA, "PromesaVenta", True)
    MsgBox "Promesa de venta guardada en:" & vbCrLf & ruta, vbInformation
    Exit Sub
Falla:
    MsgBox "No se pudo generar el PDF: " & Err.Description, vbExclamation
End Sub

Public Sub ExportarCreditoPDF()
    ' Factores + Bonos + Riesgo (uso interno, requiere clave de gerencia)
    Dim ruta As String, estaba As Boolean
    estaba = (ThisWorkbook.Worksheets(HOJA_FAC).Visible = xlSheetVisible)
    If Not estaba Then
        If Not PedirClave("Gerencia") Then Exit Sub
        Visibilidad HojasGerencia(), True
    End If
    If Not ElegirEscenarioDocumentos() Then GoTo Fin
    On Error GoTo Falla
    ruta = ExportarHojasPDF(Array(HOJA_FAC, HOJA_BON, HOJA_COMITE), "Credito", True)
    MsgBox "Expediente interno guardado en:" & vbCrLf & ruta, vbInformation
Fin:
    If Not estaba Then Visibilidad HojasGerencia(), False
    Exit Sub
Falla:
    MsgBox "No se pudo generar el PDF: " & Err.Description, vbExclamation
    Resume Fin
End Sub

Public Sub ExportarResumenComitePDF()
    Dim ruta As String, estaba As Boolean
    estaba = (ThisWorkbook.Worksheets(HOJA_COMITE).Visible = xlSheetVisible)
    If Not estaba Then
        If Not PedirClave("Gerencia") Then Exit Sub
        Visibilidad HojasGerencia(), True
    End If
    If Not ElegirEscenarioDocumentos() Then GoTo Fin
    On Error GoTo Falla
    ruta = ExportarHojaPDF(HOJA_COMITE, "Riesgo", True)
    MsgBox "Hoja de riesgo guardada en:" & vbCrLf & ruta, vbInformation
Fin:
    If Not estaba Then Visibilidad HojasGerencia(), False
    Exit Sub
Falla:
    MsgBox "No se pudo generar el PDF: " & Err.Description, vbExclamation
    Resume Fin
End Sub

Public Sub EnviarPorCorreo()
    Dim ruta As String, ol As Object, m As Object, cuerpo As String
    If Not ElegirEscenarioDocumentos() Then Exit Sub
    If Not ValidarParaEnviar() Then Exit Sub
    On Error GoTo FallaPDF
    ruta = ExportarHojasPDF(Array(HOJA_PROP, HOJA_PINI, HOJA_TABLA), "Cotizacion", False)
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
    Dim p As Variant
    p = InputBox("Plazo en meses (" & ListaPlazos() & ")" & IIf(permitirTodos, "; 0 = todos los plazos", "") & ":", titulo, _
                 IIf(permitirTodos, 0, Valor("inp_PlazoSol")))
    PedirEscenario = -1
    If p = "" Then Exit Function
    If Not IsNumeric(p) Then Exit Function
    If permitirTodos And CLng(p) = 0 Then
        PedirEscenario = 0
        Exit Function
    End If
    If EscenarioDePlazo(p) = 0 Then
        MsgBox "Ese plazo no está en la hoja Factores.", vbExclamation
        Exit Function
    End If
    PedirEscenario = EscenarioDePlazo(p)
End Function

Public Sub AplicarTasaMinima()
    Dim esc As Long, k As Long, t As Variant, resumen As String
    If Not PedirClave("Gerencia") Then Exit Sub
    esc = PedirEscenario("Aplicar tasa mínima para el cash margin objetivo", True)
    If esc < 0 Then Exit Sub
    For k = 1 To 4
        If (esc = 0 And R("esc_Incluir").Cells(k).Value = "Sí") Or esc = k Then
            t = AplicarTasaMinimaEscenario(k)
            If IsNumeric(t) Then
                resumen = resumen & R("esc_Plazo").Cells(k).Value & " meses: tasa " & Format(t, "0.00%") & " (margen sobre TIIE " & Format(t - Valor("inp_TIIE"), "0.00%") & ")" & vbCrLf
            Else
                resumen = resumen & R("esc_Plazo").Cells(k).Value & " meses: no se pudo calcular" & vbCrLf
            End If
        End If
    Next k
    Application.Calculate
    If resumen = "" Then resumen = "No se modificó ningún escenario."
    MsgBox "Margen sobre TIIE aplicado en la hoja Factores (cumple el rate card):" & _
           vbCrLf & vbCrLf & resumen, vbInformation, "Tasa mínima"
End Sub

' Escribe la tasa anual del escenario k respetando el modo de tasa (fija o TIIE + margen)
Private Sub EscribirTasa(ByVal k As Long, ByVal tasa As Double)
    If CStr(Valor("inp_ModoTasa")) = "TIIE + margen" Then
        R("esc_Margen").Cells(k).Value = tasa - CDbl(Valor("inp_TIIE"))
    Else
        R("esc_Tasa").Cells(k).Value = tasa
    End If
End Sub

' Escribe en el escenario k la tasa mínima (redondeada hacia arriba a 0.01%) y la devuelve
Public Function AplicarTasaMinimaEscenario(ByVal k As Long) As Variant
    Dim t As Variant
    Application.Calculate
    t = R("res_TasaMin").Cells(k).Value
    If IsNumeric(t) Then
        t = Application.WorksheetFunction.RoundUp(CDbl(t), 4)
        EscribirTasa k, CDbl(t)
        Application.Calculate
    End If
    AplicarTasaMinimaEscenario = t
End Function

Public Sub TasaParaRentaDeseada()
    Dim esc As Long, renta As Variant, ws As Worksheet, n As Double, M As Double, VR As Double, tipo As Long
    Dim tasa As Double
    esc = PedirEscenario("Tasa para la renta deseada", False)
    If esc < 1 Then Exit Sub
    renta = InputBox("Renta mensual deseada SIN IVA para " & R("esc_Plazo").Cells(esc).Value & " meses:", "Renta deseada")
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
              Format(tasa, "0.0000%") & "." & vbCrLf & vbCrLf & "¿Aplicarla al plazo de " & R("esc_Plazo").Cells(esc).Value & " meses?", _
              vbQuestion + vbYesNo, "Tasa para renta deseada") = vbYes Then
        EscribirTasa esc, tasa
    End If
    Exit Sub
Falla:
    MsgBox "No existe una tasa que produzca esa renta con los datos actuales.", vbExclamation
End Sub


' ---------------------------------------------------------------------
'  PROTECCIÓN (sin contraseña: sólo evita borrados accidentales)
' ---------------------------------------------------------------------
Private Function HojasProtegibles() As Variant
    HojasProtegibles = Array(HOJA_COT, HOJA_FAC, HOJA_TABLA, HOJA_PROP, HOJA_PINI, HOJA_VENTA, HOJA_BON, HOJA_COMITE, _
                             "Corrida 1", "Corrida 2", "Corrida 3", "Corrida 4")
End Function

Public Sub ProtegerHojas()
    Dim h As Variant
    For Each h In HojasProtegibles()
        ThisWorkbook.Worksheets(CStr(h)).Protect CLAVE_ADMIN, True, True, True, False, False, True, True
    Next h
    ThisWorkbook.Protect CLAVE_ADMIN, True
    MsgBox "Libro y hojas protegidos.", vbInformation
End Sub

Public Sub DesprotegerHojas()
    Dim h As Variant
    If Not PedirClave("Administrador") Then Exit Sub
    For Each h In HojasProtegibles()
        ThisWorkbook.Worksheets(CStr(h)).Unprotect CLAVE_ADMIN
    Next h
    ThisWorkbook.Unprotect CLAVE_ADMIN
    MsgBox "Libro y hojas desprotegidos. Use ""Proteger hojas"" al terminar.", vbInformation
End Sub

' ---------------------------------------------------------------------
'  NAVEGACIÓN (muestra la hoja si está oculta)
' ---------------------------------------------------------------------
Private Sub Ir(ByVal hoja As String)
    With ThisWorkbook.Worksheets(hoja)
        If .Visible <> xlSheetVisible Then
            If hoja = "Catalogos" Or Left(hoja, 7) = "Corrida" Then
                If Not PedirClave("Administrador") Then Exit Sub
            Else
                If Not PedirClave("Gerencia") Then Exit Sub
            End If
            Visibilidad Array(hoja), True
        End If
        .Activate
    End With
End Sub
Public Sub IrCotizador()
    Ir HOJA_COT
End Sub
Public Sub IrFactores()
    Ir HOJA_FAC
End Sub
Public Sub IrPropuesta()
    Ir HOJA_PROP
End Sub
Public Sub IrPagoInicial()
    Ir HOJA_PINI
End Sub
Public Sub IrVenta()
    Ir HOJA_VENTA
End Sub
Public Sub IrBonos()
    Ir HOJA_BON
End Sub
Public Sub IrRiesgo()
    Ir HOJA_COMITE
End Sub
Public Sub IrTabla()
    Ir HOJA_TABLA
End Sub
Public Sub IrHistorial()
    Ir HOJA_HIST
End Sub
Public Sub IrCatalogos()
    Ir "Catalogos"
End Sub
Public Sub IrCorrida()
    Ir "Corrida " & Valor("inp_EscTabla")
End Sub

' ---------------------------------------------------------------------
'  e-DOCUMENTOS: copia la hoja a un libro nuevo, solo valores, sin botones
' ---------------------------------------------------------------------
Private Sub GenerarArchivoCliente(ByVal hoja As String, ByVal prefijo As String)
    Dim wbN As Workbook, ws As Worksheet, i As Long, ruta As String
    If Not ValidarParaEnviar() Then Exit Sub
    On Error GoTo Falla
    Application.ScreenUpdating = False
    ThisWorkbook.Unprotect CLAVE_ADMIN
    ThisWorkbook.Worksheets(hoja).Copy
    Set wbN = ActiveWorkbook
    Set ws = wbN.Worksheets(1)
    ws.Unprotect CLAVE_ADMIN
    ws.Cells.Copy
    ws.Cells.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    For i = ws.Shapes.Count To 1 Step -1
        If Left(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
    Next i
    On Error Resume Next
    For i = wbN.Names.Count To 1 Step -1
        wbN.Names(i).Delete
    Next i
    On Error GoTo Falla
    ws.Range("A1").Select
    ws.Protect CLAVE_ADMIN
    ruta = CarpetaSalida() & Application.PathSeparator & prefijo & "_" & LimpiarNombre(CStr(Valor("inp_Folio"))) & "_" & _
           LimpiarNombre(CStr(Valor("inp_Cliente"))) & ".xlsx"
    Application.DisplayAlerts = False
    wbN.SaveAs ruta, 51
    Application.DisplayAlerts = True
    ThisWorkbook.Protect CLAVE_ADMIN, True
    Application.ScreenUpdating = True
    MsgBox "Archivo para el cliente generado (solo valores):" & vbCrLf & ruta, vbInformation, "FASTPLUS"
    Exit Sub
Falla:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error Resume Next
    ThisWorkbook.Protect CLAVE_ADMIN, True
    MsgBox "No se pudo generar el archivo: " & Err.Description, vbExclamation
End Sub

Public Sub GenerarEPropuesta()
    GenerarArchivoCliente HOJA_PROP, "e-Propuesta"
End Sub

Public Sub GenerarEPagoInicial()
    GenerarArchivoCliente HOJA_PINI, "e-PagoInicial"
End Sub

Public Sub GenerarEVenta()
    GenerarArchivoCliente HOJA_VENTA, "e-Venta"
End Sub

' ---------------------------------------------------------------------
'  RENTA DEDUCIBLE: calcula el anticipo para que la renta sea el límite
'  deducible del ISR ($200 diarios autos, $285 eléctricos/híbridos)
' ---------------------------------------------------------------------
Public Function AnticipoParaRenta(ByVal k As Long, ByVal renta As Double) As Double
    Dim ws As Worksheet, i As Double, n As Double, VR As Double, tipo As Long, M0 As Double, Mreq As Double
    Set ws = ThisWorkbook.Worksheets("Corrida " & k)
    i = ws.Range("C9").Value
    n = ws.Range("C5").Value
    VR = ws.Range("C21").Value
    tipo = ws.Range("C7").Value
    M0 = ws.Range("C19").Value + ws.Range("C16").Value          ' monto a financiar sin anticipo
    Mreq = -Application.WorksheetFunction.PV(i, n, renta, VR, tipo)   ' monto que produce esa renta
    AnticipoParaRenta = M0 - Mreq
End Function

Public Sub AjustarRentaDeducible()
    Dim k As Long, lim As Double, a As Double, tipoAct As String
    k = CLng(Valor("inp_EscTabla"))
    tipoAct = CStr(Valor("inp_TipoActivo"))
    If tipoAct = "Automóvil eléctrico / híbrido" Then
        lim = CDbl(Valor("par_LimAutoEV")) * 30
    Else
        lim = CDbl(Valor("par_LimAuto")) * 30
        If tipoAct <> "Automóvil" Then
            If MsgBox("El límite de deducibilidad aplica a automóviles. ¿Ajustar de todos modos la renta a " & _
                      Format(lim, "$#,##0.00") & "?", vbQuestion + vbYesNo, "Renta deducible") <> vbYes Then Exit Sub
        End If
    End If
    a = AnticipoParaRenta(k, lim)
    If a <= 0 Then
        MsgBox "La renta de " & Valor("inp_PlazoSol") & " meses ya es menor o igual a " & Format(lim, "$#,##0.00") & _
               "; no se requiere anticipo adicional (deducible al 100%).", vbInformation, "Renta deducible"
        Exit Sub
    End If
    If a >= CDbl(Valor("inp_Valor")) Then
        MsgBox "No es posible: el anticipo necesario supera el valor del equipo.", vbExclamation, "Renta deducible"
        Exit Sub
    End If
    If MsgBox("Para que la renta mensual de " & Valor("inp_PlazoSol") & " meses sea " & Format(lim, "$#,##0.00") & _
              " (100% deducible) se requiere una renta extraordinaria de " & Format(a, "$#,##0.00") & " + IVA." & _
              vbCrLf & vbCrLf & "¿Aplicarla?", vbQuestion + vbYesNo, "Renta deducible") <> vbYes Then Exit Sub
    R("inp_Anticipo").Value = Round(a, 2)
    Application.Calculate
    MsgBox "Renta mensual: " & Format(R("res_Renta").Cells(k).Value, "$#,##0.00") & " + IVA", vbInformation, "Renta deducible"
End Sub

' ---------------------------------------------------------------------
'  BONOS: mostrar / ocultar la distribución (como ABC)
' ---------------------------------------------------------------------
Public Sub MostrarBonos()
    If Not PedirClave("Gerencia") Then Exit Sub
    With ThisWorkbook.Worksheets(HOJA_BON)
        .Unprotect CLAVE_ADMIN
        R("bon_Distribucion").EntireRow.Hidden = False
        .Protect CLAVE_ADMIN
    End With
End Sub

Public Sub OcultarBonos()
    With ThisWorkbook.Worksheets(HOJA_BON)
        .Unprotect CLAVE_ADMIN
        R("bon_Distribucion").EntireRow.Hidden = True
        .Protect CLAVE_ADMIN
    End With
End Sub
