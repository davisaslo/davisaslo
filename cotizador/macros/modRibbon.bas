Option Explicit

' Llamadas desde la pestaña "Cotizador Arrendamiento" de la cinta de opciones

Public Sub rbNueva(control As IRibbonControl)
    NuevaCotizacion
End Sub

Public Sub rbGuardar(control As IRibbonControl)
    GuardarCotizacion
End Sub

Public Sub rbCargar(control As IRibbonControl)
    CargarCotizacion
End Sub

Public Sub rbPaquete(control As IRibbonControl)
    ExportarPaqueteCliente
End Sub

Public Sub rbPropuesta(control As IRibbonControl)
    ExportarPropuestaPDF
End Sub

Public Sub rbCartaPDF(control As IRibbonControl)
    ExportarCartaPDF
End Sub

Public Sub rbTablaPDF(control As IRibbonControl)
    ExportarTablaPDF
End Sub

Public Sub rbPromesa(control As IRibbonControl)
    ExportarPromesaVentaPDF
End Sub

Public Sub rbComite(control As IRibbonControl)
    ExportarResumenComitePDF
End Sub

Public Sub rbCorreo(control As IRibbonControl)
    EnviarPorCorreo
End Sub

Public Sub rbTasaMin(control As IRibbonControl)
    AplicarTasaMinima
End Sub

Public Sub rbTasaRenta(control As IRibbonControl)
    TasaParaRentaDeseada
End Sub

Public Sub rbCopiar(control As IRibbonControl)
    CopiarEscenario1
End Sub

Public Sub rbProteger(control As IRibbonControl)
    ProtegerHojas
End Sub

Public Sub rbDesproteger(control As IRibbonControl)
    DesprotegerHojas
End Sub

Public Sub rbIrCotizador(control As IRibbonControl)
    IrCotizador
End Sub

Public Sub rbIrPropuesta(control As IRibbonControl)
    IrPropuesta
End Sub

Public Sub rbIrCarta(control As IRibbonControl)
    IrCarta
End Sub

Public Sub rbIrPagoInicial(control As IRibbonControl)
    IrPagoInicial
End Sub

Public Sub rbIrTabla(control As IRibbonControl)
    IrTabla
End Sub

Public Sub rbIrPromesa(control As IRibbonControl)
    IrPromesa
End Sub

Public Sub rbIrComite(control As IRibbonControl)
    IrComite
End Sub

Public Sub rbIrHistorial(control As IRibbonControl)
    IrHistorial
End Sub

Public Sub rbIrSensibilidad(control As IRibbonControl)
    IrSensibilidad
End Sub

Public Sub rbIrCorrida(control As IRibbonControl)
    IrCorrida
End Sub

Public Sub rbIrConfig(control As IRibbonControl)
    IrConfiguracion
End Sub

Public Sub rbIrInicio(control As IRibbonControl)
    IrInicio
End Sub
