Option Explicit

' Llamadas desde la pestaña "FASTPLUS" de la cinta de opciones

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

Public Sub rbTablaPDF(control As IRibbonControl)
    ExportarTablaPDF
End Sub

Public Sub rbVenta(control As IRibbonControl)
    ExportarPromesaVentaPDF
End Sub

Public Sub rbCorreo(control As IRibbonControl)
    EnviarPorCorreo
End Sub

Public Sub rbCredito(control As IRibbonControl)
    ExportarCreditoPDF
End Sub

Public Sub rbRiesgo(control As IRibbonControl)
    ExportarResumenComitePDF
End Sub

Public Sub rbTasaMin(control As IRibbonControl)
    AplicarTasaMinima
End Sub

Public Sub rbTasaRenta(control As IRibbonControl)
    TasaParaRentaDeseada
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

Public Sub rbIrFactores(control As IRibbonControl)
    IrFactores
End Sub

Public Sub rbIrPropuesta(control As IRibbonControl)
    IrPropuesta
End Sub

Public Sub rbIrPagoInicial(control As IRibbonControl)
    IrPagoInicial
End Sub

Public Sub rbIrVenta(control As IRibbonControl)
    IrVenta
End Sub

Public Sub rbIrBonos(control As IRibbonControl)
    IrBonos
End Sub

Public Sub rbIrRiesgo(control As IRibbonControl)
    IrRiesgo
End Sub

Public Sub rbIrTabla(control As IRibbonControl)
    IrTabla
End Sub

Public Sub rbIrHistorial(control As IRibbonControl)
    IrHistorial
End Sub

Public Sub rbIrCatalogos(control As IRibbonControl)
    IrCatalogos
End Sub

Public Sub rbIrCorrida(control As IRibbonControl)
    IrCorrida
End Sub

Public Sub rbEPropuesta(control As IRibbonControl)
    GenerarEPropuesta
End Sub

Public Sub rbEPago(control As IRibbonControl)
    GenerarEPagoInicial
End Sub

Public Sub rbEVenta(control As IRibbonControl)
    GenerarEVenta
End Sub

Public Sub rbRentaDed(control As IRibbonControl)
    AjustarRentaDeducible
End Sub

Public Sub rbGerencia(control As IRibbonControl)
    AccesoGerencia
End Sub

Public Sub rbOcultarGer(control As IRibbonControl)
    OcultarGerencia
End Sub

Public Sub rbAdmin(control As IRibbonControl)
    AccesoAdministrador
End Sub
