Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    Application.OnKey "^+n", "'" & ThisWorkbook.Name & "'!NuevaCotizacion"
    Application.OnKey "^+g", "'" & ThisWorkbook.Name & "'!GuardarCotizacion"
    Application.OnKey "^+p", "'" & ThisWorkbook.Name & "'!ExportarPaqueteCliente"
    OcultarRestringidas
    ThisWorkbook.Worksheets("Cotizador").Activate
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.OnKey "^+n"
    Application.OnKey "^+g"
    Application.OnKey "^+p"
End Sub
