# Cotizador Master FASTPLUS

Archivo: **`Cotizador_FASTPLUS.xlsm`**. Tiene el mismo formato y el mismo flujo de trabajo que el
cotizador de ABC Leasing (PROCAAR), con la marca FASTPLUS / SOFOPLUS y un motor de cálculo propio
verificado.

> La versión de este repositorio público trae promotores de ejemplo. La versión con el catálogo
> real de colaboradores se entregó por separado.

## Hojas (mismo orden que ABC)

| Hoja | Qué hace |
|---|---|
| **Cotizador** | Captura con la misma distribución que ABC: datos generales, datos para el cálculo (IVA incluido, excepto anticipo), seguro y tipo de cambio. Abajo, la tabla de condiciones para 12 / 24 / 36 / 48 meses, con el plazo solicitado resaltado. |
| **Factores** | Factores por plazo: margen sobre TIIE, % residual, comisión y depósito. También pagos mensuales, residual, deducibilidad, análisis de rentabilidad contra el rate card con observaciones en rojo, análisis por plazo y el rate card por TIER de ventas. |
| **Propuesta** | Propuesta para el cliente con el formato de ABC: datos de la oficina a la izquierda y logo a la derecha, valor y conceptos financiados, pago inicial, pagos mensuales, valor residual y vigencia. |
| **Pago Inicial** | Conceptos del pago inicial y datos para depósito (transferencia, sucursal y beneficiario). |
| **Venta** | Carta de promesa de venta del bien al término del plazo, por el valor residual. |
| **Bonos** | Margen generado y su distribución por centro de costos. |
| **Riesgo** | Hoja de datos para riesgo con TIR, margen, rate card, dictamen, datos del acta y autorizaciones. |
| Tabla de Pagos | Calendario de pagos con IVA del plazo solicitado. |
| Historial | Folios guardados; permite recargarlos. |
| Catalogos | Empresa, IVA, TIIE, fondeo, banco, oficinas por región y catálogo de promotores. |
| Corrida 1–4 (ocultas) | Corrida completa de cada plazo. Se ven con el botón *Ir a → Corrida*. |

## Uso

1. Abra el archivo y pulse **Habilitar contenido**. Aparece la pestaña **FASTPLUS** en la cinta.
2. En **Cotizador** capture las celdas blancas: promotor, cliente, activo, precio con IVA, anticipo,
   TIIE, plazo solicitado, seguro, GPS y otros gastos.
3. Revise la tabla de condiciones y el dictamen. Si un plazo no cumple, use **Margen mínimo (rate
   card)**.
4. **Propuesta completa PDF** (Ctrl+Shift+P) genera Propuesta + Pago Inicial + Tabla de Pagos.
   También puede usar **Enviar por correo**.
5. Para crédito use **Expediente interno PDF**, que incluye Factores + Bonos + Riesgo.

## Mejoras sobre el archivo de ABC

* **Plazos:** los 4 plazos son editables (1 a 84 meses). En ABC, el plazo de 60 meses da `#NUM!`.
* **TIR y margen de caja:** calculados con fechas reales y costo de fondeo. En el archivo de ejemplo
  de ABC la TIR sale con error y el margen en −1014%.
* **Margen mínimo sobre TIIE:** se calcula en forma exacta para cumplir el rate card y se aplica
  con un botón.
* **Tasa para la renta deseada:** calcula la tasa que da la renta que pide el cliente.
* **Desglose exacto de la renta** en equipo, GPS, seguro y otros. La suma cuadra al centavo.
* **Comprobación automática** de que cada corrida termina en cero.
* **Historial, paquete PDF y correo** por Outlook.
* **Mismo resultado que ABC con los mismos datos:** la renta base coincide al centavo.

## Pendiente de capturar (hoja Catalogos / Factores)

* Banco, CLABE, cuenta y convenio para el pago inicial.
* Domicilio y teléfono de las oficinas de Guadalajara, Cancún y Querétaro.
* TIIE vigente y tasa de fondeo.
* **Factores y rate card reales de FASTPLUS.** Los valores actuales son sugeridos, no son los de ABC.
* Logo de FASTPLUS. Hoy se usa el de SOFOPLUS: clic derecho en la imagen → *Cambiar imagen*.

## Verificación

* 8,443 fórmulas sin errores.
* Renta, monto financiado, depósito, comisión, margen neto, TIR, tasa mínima y pago inicial se
  compararon contra un modelo independiente en Python, en dos casos. El segundo incluye pago
  anticipado, renta proporcional, descuento del proveedor, seguro, GPS y otros financiados,
  depósito en % y comisión del originador. Coinciden al centavo.
* Las macros se probaron en LibreOffice: guardar, limpiar, recargar y margen mínimo. **No se
  probaron en Microsoft Excel.** Si una macro falla, importe los módulos de `macros/`.
