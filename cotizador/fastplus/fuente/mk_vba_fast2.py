"""Segunda capa: seguridad, botones en hojas, e-documentos y renta deducible (estilo ABC)."""
import re
import subprocess

subprocess.run(['python3', 'mk_vba_fast.py'], check=True)
import os
CLAVES = {'CLAVE_ADMIN': os.environ.get('CLAVE_ADMIN', 'FP-Admin2026'), 'CLAVE_GERENCIA': os.environ.get('CLAVE_GERENCIA', 'FP-Gerencia2026')}
s = open('vba_fast/modCotizador.bas', encoding='utf-8').read()


def rep(a, b):
    global s
    assert a in s, a[:80]
    s = s.replace(a, b, 1)


def span(start, end_marker):
    i = s.index(start)
    j = s.index(end_marker, i) + len(end_marker)
    return i, j


rep('Private Const COL_SNAPSHOT As Long = 22', '''Private Const CLAVE_ADMIN As String = "%s"        ' protege libro y hojas; abre Catálogos
Private Const CLAVE_GERENCIA As String = "%s"     ' abre Factores, Bonos y Riesgo
Private Const COL_SNAPSHOT As Long = 22''' % (CLAVES['CLAVE_ADMIN'], CLAVES['CLAVE_GERENCIA']))

# ---------------- seguridad y visibilidad
rep("' ---------------------------------------------------------------------\n'  NUEVA COTIZACIÓN", '''' ---------------------------------------------------------------------
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
'  NUEVA COTIZACIÓN''')

# ---------------- PDF de hojas restringidas: muestra temporalmente
rep('''Public Sub ExportarCreditoPDF()
    ' Factores + Bonos + Riesgo (uso interno)
    Dim ruta As String
    If Not ElegirEscenarioDocumentos() Then Exit Sub
    On Error GoTo Falla
    ruta = ExportarHojasPDF(Array(HOJA_FAC, HOJA_BON, HOJA_COMITE), "Credito", True)''', '''Public Sub ExportarCreditoPDF()
    ' Factores + Bonos + Riesgo (uso interno, requiere clave de gerencia)
    Dim ruta As String, estaba As Boolean
    estaba = (ThisWorkbook.Worksheets(HOJA_FAC).Visible = xlSheetVisible)
    If Not estaba Then
        If Not PedirClave("Gerencia") Then Exit Sub
        Visibilidad HojasGerencia(), True
    End If
    If Not ElegirEscenarioDocumentos() Then GoTo Fin
    On Error GoTo Falla
    ruta = ExportarHojasPDF(Array(HOJA_FAC, HOJA_BON, HOJA_COMITE), "Credito", True)''')
i, j = span("    MsgBox \"Expediente interno guardado en:\" & vbCrLf & ruta, vbInformation", 'End Sub')
s = s[:i] + '''    MsgBox "Expediente interno guardado en:" & vbCrLf & ruta, vbInformation
Fin:
    If Not estaba Then Visibilidad HojasGerencia(), False
    Exit Sub
Falla:
    MsgBox "No se pudo generar el PDF: " & Err.Description, vbExclamation
    Resume Fin
End Sub''' + s[j:]
i, j = span('Public Sub ExportarResumenComitePDF()', 'End Sub')
s = s[:i] + '''Public Sub ExportarResumenComitePDF()
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
End Sub''' + s[j:]

# ---------------- margen mínimo: requiere gerencia (cambia política)
rep('''Public Sub AplicarTasaMinima()
    Dim esc As Long, k As Long, t As Variant, resumen As String''', '''Public Sub AplicarTasaMinima()
    Dim esc As Long, k As Long, t As Variant, resumen As String
    If Not PedirClave("Gerencia") Then Exit Sub''')

# ---------------- proteger / desproteger con clave
i, j = span('Public Sub ProtegerHojas()', 'End Sub')
s = s[:i] + '''Public Sub ProtegerHojas()
    Dim h As Variant
    For Each h In HojasProtegibles()
        ThisWorkbook.Worksheets(CStr(h)).Protect CLAVE_ADMIN, True, True, True, False, False, True, True
    Next h
    ThisWorkbook.Protect CLAVE_ADMIN, True
    MsgBox "Libro y hojas protegidos.", vbInformation
End Sub''' + s[j:]
i, j = span('Public Sub DesprotegerHojas()', 'End Sub')
s = s[:i] + '''Public Sub DesprotegerHojas()
    Dim h As Variant
    If Not PedirClave("Administrador") Then Exit Sub
    For Each h In HojasProtegibles()
        ThisWorkbook.Worksheets(CStr(h)).Unprotect CLAVE_ADMIN
    Next h
    ThisWorkbook.Unprotect CLAVE_ADMIN
    MsgBox "Libro y hojas desprotegidos. Use ""Proteger hojas"" al terminar.", vbInformation
End Sub''' + s[j:]

# ---------------- navegación: hojas restringidas piden clave
rep('''Private Sub Ir(ByVal hoja As String)
    With ThisWorkbook.Worksheets(hoja)
        If .Visible <> xlSheetVisible Then .Visible = xlSheetVisible
        .Activate
    End With
End Sub''', '''Private Sub Ir(ByVal hoja As String)
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
End Sub''')

# ---------------- e-documentos, renta deducible, bonos
s = s.rstrip('\n') + '''

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
'''
s = s.replace('ThisWorkbook.Worksheets(CStr(h)).Unprotect\n', 'ThisWorkbook.Worksheets(CStr(h)).Unprotect CLAVE_ADMIN\n')
rep('    LimpiarNombre = Replace(s, " ", "_")', '    s = Replace(Replace(s, ".", ""), ",", "")\n    LimpiarNombre = Replace(s, " ", "_")')
open('vba_fast/modCotizador.bas', 'w', encoding='utf-8').write(s)

# ThisWorkbook: ocultar restringidas al abrir
t = open('vba_fast/ThisWorkbook.cls', encoding='utf-8').read()
t = t.replace('''    ThisWorkbook.Worksheets("Cotizador").Activate
End Sub''', '''    OcultarRestringidas
    ThisWorkbook.Worksheets("Cotizador").Activate
End Sub''', 1)
assert 'OcultarRestringidas' in t
open('vba_fast/ThisWorkbook.cls', 'w', encoding='utf-8').write(t)

# cinta: nuevas acciones
extra = [('rbEPropuesta', 'GenerarEPropuesta'), ('rbEPago', 'GenerarEPagoInicial'), ('rbEVenta', 'GenerarEVenta'),
         ('rbRentaDed', 'AjustarRentaDeducible'), ('rbGerencia', 'AccesoGerencia'), ('rbOcultarGer', 'OcultarGerencia'),
         ('rbAdmin', 'AccesoAdministrador')]
rb = open('vba_fast/modRibbon.bas', encoding='utf-8').read().rstrip('\n')
for a, b in extra:
    rb += '\n\nPublic Sub %s(control As IRibbonControl)\n    %s\nEnd Sub' % (a, b)
open('vba_fast/modRibbon.bas', 'w', encoding='utf-8').write(rb + '\n')
print('ok', CLAVES)
