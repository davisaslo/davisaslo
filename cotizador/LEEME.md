# Cotizador de Arrendamiento – SOFOPLUS (versión 3)

Archivo: **`Cotizador_Arrendamiento_SOFOPLUS.xlsm`** (Excel con macros).

> La versión de este repositorio público trae **promotores de ejemplo**. La versión con la lista
> real de colaboradores se entregó por separado; para cargarla en esta, capture los promotores en
> la hoja *Configuracion* (tabla "PROMOTORES").

## Comparativo contra el cotizador de ABC Leasing (PROCAAR 14.61)

| Función | ABC Leasing | SOFOPLUS v3 |
|---|---|---|
| Tasa = TIIE + margen por plazo | Sí | Sí (o tasa fija, a elegir) |
| Precio capturado con IVA | Sí | Sí (con o sin IVA) |
| Anticipo como renta extraordinaria | Sí | Sí (o como enganche) |
| Seguro, GPS, mantenimiento y garantía financiados o de contado | Sí | Sí |
| Comparativo de varios plazos | 4 plazos fijos (60 m da `#NUM!`) | 4 escenarios libres de 1 a 84 meses + sucesivo |
| Propuesta ejecutiva de un plazo | Sí | Sí, con desglose exacto de la renta |
| Hoja de pago inicial con datos bancarios | Sí | Sí, la referencia es el folio |
| Carta de promesa de venta del residual | Sí (la fecha de pago sale errónea) | Sí, con fecha real del último pago |
| Rate card / mínimos de rentabilidad | Sí (TIR y margen con error en el ejemplo) | CM neto, TIR mínima y residual máximo por plazo |
| Margen después de comisión del promotor | Sí | Sí |
| Deducibilidad de la renta (autos) | Sí | Sí: autos $200 diarios, eléctricos/híbridos $285 |
| Hoja para comité de crédito | Sí ("Riesgo" / "Lease Economics") | Sí, con dictamen, alertas y firmas |
| Tasa mínima para cumplir el margen | No | Sí, exacta y aplicable con un botón |
| Tasa para la renta que pide el cliente | No | Sí |
| Tabla de amortización que cuadra en cero | No | Sí, con verificación automática |
| Historial y recarga de cotizaciones | No | Sí |
| Paquete PDF y envío por Outlook | No | Sí |
| Pestaña propia en la cinta de Excel | No | Sí |

Con los mismos datos que su archivo de ejemplo, la renta base de SOFOPLUS coincide al centavo con
la de ABC (24 meses: $40,084.87).

## Cómo empezar

1. Descargue el archivo. Si Windows lo bloquea: clic derecho → *Propiedades* → **Desbloquear**.
2. Ábralo y pulse **Habilitar contenido**.
3. Use la pestaña **"Cotizador Arrendamiento"** de la cinta (o Alt+F8).
4. Capture solo las **celdas amarillas** de la hoja **Cotizador**.

| Atajo | Acción |
|---|---|
| Ctrl+Shift+N | Nueva cotización |
| Ctrl+Shift+G | Guardar en Historial |
| Ctrl+Shift+P | Paquete completo en PDF |

## Hojas

| Hoja | Uso |
|---|---|
| Inicio | Guía, accesos, macros y modelo matemático |
| Cotizador | Captura, resultados por escenario, rentabilidad vs. rate card y alertas |
| Propuesta | Propuesta ejecutiva del escenario elegido (formato tipo ABC, marca SOFOPLUS) |
| Carta Cotizacion | Carta comparativa de hasta 4 plazos con términos legales |
| Pago Inicial | Conceptos del pago inicial y datos para el depósito |
| Tabla de Pagos | Calendario de pagos con IVA |
| Promesa de Venta | Carta de venta del equipo al término del arrendamiento |
| Resumen Comite | Hoja interna para comité de crédito (no se entrega al cliente) |
| Sensibilidad | Rentas por plazo (12–84 m) × tasa |
| Historial | Registro de folios; permite recargarlos |
| Corrida 1–4 | Corridas completas: amortización, flujos, VP, TIR, CAT |
| Configuracion | Empresa, IVA, TIIE, fondeo, rate card, oficinas, promotores, banco y textos |

## Pendientes de capturar en *Configuracion*

* **Datos bancarios:** banco, cuenta, CLABE y convenio para el pago inicial.
* **Domicilios y teléfonos** de las oficinas de Guadalajara, Cancún y Querétaro. Mientras estén
  vacíos se usa el domicilio corporativo.
* **TIIE vigente y tasa de fondeo.**
* **Rate card:** CM mínimo, TIR mínima y residual máximo por plazo. Los valores iniciales son
  sugeridos (CM 8%, TIR 23%, residual 40% a 20%).
* **Comisión del promotor** y gastos de investigación por defecto.

## Modelo matemático

* Renta: R = PMT(i, n, −M, VR, tipo), con i = tasa/12 (tasa fija o TIIE + margen),
  M = valor − anticipo + financiados y VR = % residual × valor.
* Desglose exacto: R = PMT(equipo) + PMT(seguro) + PMT(GPS y otros), por linealidad de PMT.
* Amortización: interés = (saldo − tipo × pago) × i. La tabla termina en 0.
* Cash margin: VP al fondeo de todos los flujos, incluido el costo diario por los días entre la
  firma y el inicio. **Neto** = CM − comisión del promotor.
* Tasa mínima: el CM es lineal en la renta, así que R_mín = (CM objetivo × M + comisión − A)/B y
  la tasa = RATE(n, R_mín, −M, VR, tipo) × 12. En modo TIIE se aplica como margen.
* TIR (XIRR) con fechas reales; CAT informativo sin IVA.

## Verificación realizada

* 8,640 fórmulas sin errores.
* Renta, desglose, cash margin neto, TIR, CAT, tasa mínima, pago inicial y deducibilidad
  coinciden al centavo con un modelo independiente en Python. Se probó en dos configuraciones:
  la del archivo original y la del ejemplo de ABC (TIIE + margen, precio con IVA, anticipo 10%,
  comisión 2%, pago vencido, otros financiados).
* Macros probadas de punta a punta en LibreOffice: guardar, limpiar, recargar, tasa mínima en modo
  fijo y en modo TIIE + margen, y búsqueda de los datos del promotor.
* **No se probó en Microsoft Excel.** Si alguna macro falla, importe los módulos de `macros/`
  (Alt+F11 → Archivo → Importar).

## Hallazgos del archivo original (v1, corregidos)

La renta usaba un residual de 5% y el plazo sucesivo cobraba 40%, lo que inflaba la TIR (52%) y el
cash margin (24%). Con tasa 17% y fondeo 20.5% el spread es negativo. Detalle en el historial de
versiones del repositorio.
