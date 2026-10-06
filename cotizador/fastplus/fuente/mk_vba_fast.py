import os
s = open('vba/modCotizador.bas', encoding='utf-8').read()


def rep(a, b):
    global s
    assert a in s, a[:80]
    s = s.replace(a, b, 1)


def span(start, end_marker):
    i = s.index(start)
    j = s.index(end_marker, i)
    return i, j


s = s.replace("'  COTIZADOR DE ARRENDAMIENTO - Macros", "'  COTIZADOR MASTER FASTPLUS - Macros").replace('par_*, def_*, res_*', 'par_*, rc_*, res_*')
rep('Private Const HOJA_CARTA As String = "Carta Cotizacion"\n',
    'Private Const HOJA_FAC As String = "Factores"\nPrivate Const HOJA_BON As String = "Bonos"\n')
rep('Private Const HOJA_VENTA As String = "Promesa de Venta"\nPrivate Const HOJA_COMITE As String = "Resumen Comite"',
    'Private Const HOJA_VENTA As String = "Venta"\nPrivate Const HOJA_COMITE As String = "Riesgo"')

i, j = span('Private Function NombresEntrada() As Variant', 'End Function')
s = s[:i] + '''Private Function NombresEntrada() As Variant
    NombresEntrada = Array("inp_Folio", "inp_Fecha", "inp_Cliente", "inp_RFC", "inp_Contacto", "inp_Correo", _
        "inp_Proveedor", "inp_Equipo", "inp_Promotor", "inp_Producto", "inp_Obligado", "inp_Tipo", "inp_Moneda", _
        "inp_TC", "inp_Modalidad", "inp_TipoActivo", "inp_Estado", "inp_Originador", "inp_Tier", "inp_Precio", _
        "inp_Descuento", "inp_DiasRP", "inp_IVAEquipo", "inp_UsoAnticipo", "inp_Anticipo", "inp_SeguroFin", _
        "inp_TIIE", "inp_PlazoSol", "inp_Fuente", "inp_FondeoManual", "inp_OtrosMonto", "inp_OtrosDesc", "inp_GPSMonto", _
        "inp_GastosInv", "inp_ComBanco", "inp_CMobj", "inp_ComProm", "inp_SeguroMonto", "inp_Aseguradora", _
        "inp_AseguradoraPor", "esc_Plazo", "esc_Margen", "esc_Residual", "esc_Comision", "esc_DepPct", _
        "esc_Deposito", "esc_Incluir")
''' + s[j:]

i, j = span('Public Sub LimpiarCaptura(', 'End Sub')
s = s[:i] + '''Public Sub LimpiarCaptura(Optional ByVal sinDialogo As Boolean = True)
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
''' + s[j:]

s = s.replace('.Cells(1, k)', '.Cells(k)').replace('.Cells(1, esc)', '.Cells(esc)')
rep('''    col = COL_SNAPSHOT
    For Each nm In NombresEntrada()
        For Each c In R(CStr(nm)).Cells
            If c.MergeArea.Cells(1, 1).Address = c.Address Then
                ws.Cells(FILA_ENCABEZADO, col).Value = CStr(nm) & "|" & (c.Column - R(CStr(nm)).Column + 1)
                ws.Cells(fila, col).Value = c.Value
                col = col + 1
            End If
        Next c
    Next nm''', '''    col = COL_SNAPSHOT
    For Each nm In NombresEntrada()
        idx = 0
        For Each c In R(CStr(nm)).Cells
            idx = idx + 1
            ws.Cells(FILA_ENCABEZADO, col).Value = CStr(nm) & "|" & idx
            ws.Cells(fila, col).Value = c.Value
            col = col + 1
        Next c
    Next nm''')
rep('    Dim col As Long, k As Long, plazos As String, esNuevo As Boolean',
    '    Dim col As Long, k As Long, plazos As String, esNuevo As Boolean, idx As Long')
rep('Set destino = R(partes(0)).Cells(1, CLng(partes(1)))', 'Set destino = R(partes(0)).Cells(CLng(partes(1)))')

i, j = span('Private Function ElegirEscenarioDocumentos() As Boolean', 'End Function')
s = s[:i] + '''Private Function EscenarioDePlazo(ByVal plazo As Variant) As Long
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
''' + s[j:]

i, j = span('Private Function ValidarParaEnviar() As Boolean', 'End Function')
s = s[:i] + '''Private Function ValidarParaEnviar() As Boolean
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
''' + s[j:]

i, j = span('Public Sub ExportarCartaPDF()', 'Public Sub ExportarTablaPDF()')
s = s[:i] + s[j:]
s = s.replace('Array(HOJA_PROP, HOJA_CARTA, HOJA_PINI, HOJA_TABLA)', 'Array(HOJA_PROP, HOJA_PINI, HOJA_TABLA)')
s = s.replace("' Propuesta + carta comparativa + instrucciones de pago + tabla de pagos en un solo PDF",
              "' Propuesta + pago inicial + tabla de pagos en un solo PDF")
rep('Public Sub ExportarResumenComitePDF()', '''Public Sub ExportarCreditoPDF()
    ' Factores + Bonos + Riesgo (uso interno)
    Dim ruta As String
    If Not ElegirEscenarioDocumentos() Then Exit Sub
    On Error GoTo Falla
    ruta = ExportarHojasPDF(Array(HOJA_FAC, HOJA_BON, HOJA_COMITE), "Credito", True)
    MsgBox "Expediente interno guardado en:" & vbCrLf & ruta, vbInformation
    Exit Sub
Falla:
    MsgBox "No se pudo generar el PDF: " & Err.Description, vbExclamation
End Sub

Public Sub ExportarResumenComitePDF()''')
s = s.replace('ruta = ExportarHojaPDF(HOJA_COMITE, "Comite", True)', 'ruta = ExportarHojaPDF(HOJA_COMITE, "Riesgo", True)')
s = s.replace('"Resumen para comité guardado en:"', '"Hoja de riesgo guardada en:"')

i, j = span('Private Function PedirEscenario(', 'End Function')
s = s[:i] + '''Private Function PedirEscenario(ByVal titulo As String, ByVal permitirTodos As Boolean) As Long
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
''' + s[j:]
rep('resumen = resumen & "Escenario " & k & ": " & Format(t, "0.00%") & vbCrLf',
    'resumen = resumen & R("esc_Plazo").Cells(k).Value & " meses: tasa " & Format(t, "0.00%") & " (margen sobre TIIE " & Format(t - Valor("inp_TIIE"), "0.00%") & ")" & vbCrLf')
rep('resumen = resumen & "Escenario " & k & ": no se pudo calcular" & vbCrLf',
    'resumen = resumen & R("esc_Plazo").Cells(k).Value & " meses: no se pudo calcular" & vbCrLf')
rep('MsgBox "Tasa anual aplicada (cash margin objetivo " & Format(Valor("inp_CMobj"), "0.00%") & "):" & _',
    'MsgBox "Margen sobre TIIE aplicado en la hoja Factores (cumple el rate card):" & _')
rep('"¿Aplicarla al escenario " & esc & "?"', '"¿Aplicarla al plazo de " & R("esc_Plazo").Cells(esc).Value & " meses?"')
rep('renta = InputBox("Renta mensual deseada SIN IVA para el escenario " & esc & ":", "Renta deseada")',
    'renta = InputBox("Renta mensual deseada SIN IVA para " & R("esc_Plazo").Cells(esc).Value & " meses:", "Renta deseada")')
i, j = span('Public Sub CopiarEscenario1()', 'End Sub')
s = s[:i] + s[j + len('End Sub\n'):]
rep('    HojasProtegibles = Array(HOJA_COT, HOJA_CARTA, HOJA_TABLA, "Sensibilidad", HOJA_PROP, HOJA_PINI, HOJA_VENTA, HOJA_COMITE, _',
    '    HojasProtegibles = Array(HOJA_COT, HOJA_FAC, HOJA_TABLA, HOJA_PROP, HOJA_PINI, HOJA_VENTA, HOJA_BON, HOJA_COMITE, _')
i = s.index("' ---------------------------------------------------------------------\n'  NAVEGACIÓN")
nav = [('IrCotizador', 'HOJA_COT'), ('IrFactores', 'HOJA_FAC'), ('IrPropuesta', 'HOJA_PROP'), ('IrPagoInicial', 'HOJA_PINI'),
       ('IrVenta', 'HOJA_VENTA'), ('IrBonos', 'HOJA_BON'), ('IrRiesgo', 'HOJA_COMITE'), ('IrTabla', 'HOJA_TABLA'),
       ('IrHistorial', 'HOJA_HIST'), ('IrCatalogos', '"Catalogos"'), ('IrCorrida', '"Corrida " & Valor("inp_EscTabla")')]
s = s[:i] + '''' ---------------------------------------------------------------------
'  NAVEGACIÓN (muestra la hoja si está oculta)
' ---------------------------------------------------------------------
Private Sub Ir(ByVal hoja As String)
    With ThisWorkbook.Worksheets(hoja)
        If .Visible <> xlSheetVisible Then .Visible = xlSheetVisible
        .Activate
    End With
End Sub
''' + ''.join('Public Sub %s()\n    Ir %s\nEnd Sub\n' % (a, b) for a, b in nav)
for bad in ('HOJA_CARTA', 'Sensibilidad', 'def_', 'Cells(1, k)'):
    assert bad not in s, bad
os.makedirs('vba_fast', exist_ok=True)
open('vba_fast/modCotizador.bas', 'w', encoding='utf-8').write(s)

pairs = [('rbNueva', 'NuevaCotizacion'), ('rbGuardar', 'GuardarCotizacion'), ('rbCargar', 'CargarCotizacion'),
         ('rbPaquete', 'ExportarPaqueteCliente'), ('rbPropuesta', 'ExportarPropuestaPDF'), ('rbTablaPDF', 'ExportarTablaPDF'),
         ('rbVenta', 'ExportarPromesaVentaPDF'), ('rbCorreo', 'EnviarPorCorreo'), ('rbCredito', 'ExportarCreditoPDF'),
         ('rbRiesgo', 'ExportarResumenComitePDF'), ('rbTasaMin', 'AplicarTasaMinima'), ('rbTasaRenta', 'TasaParaRentaDeseada'),
         ('rbProteger', 'ProtegerHojas'), ('rbDesproteger', 'DesprotegerHojas')] + [('rb' + a, a) for a, _ in nav]
out = ["Option Explicit", "", "' Llamadas desde la pestaña \"FASTPLUS\" de la cinta de opciones", ""]
for a, b in pairs:
    out += ["Public Sub %s(control As IRibbonControl)" % a, "    %s" % b, "End Sub", ""]
open('vba_fast/modRibbon.bas', 'w', encoding='utf-8').write('\n'.join(out))
open('vba_fast/ThisWorkbook.cls', 'w', encoding='utf-8').write(open('vba/ThisWorkbook.cls', encoding='utf-8').read())
print('ok')
