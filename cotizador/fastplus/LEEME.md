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

## Política de tasas y fondeo (versión 4)

* **Fondeo SOFOPLUS / FASTPLUS: 21.0%** (tasa fija). Es la fuente por defecto en el catálogo.
* **Tasa a clientes: del 26% al 33% anual** (parámetros *Tasa mínima / máxima a clientes* en
  Catálogos).
  * Una tasa fuera de rango se marca en rojo en Factores, el dictamen dice **FUERA DE RANGO** y
    aparece una alerta.
  * La "tasa mínima para cumplir" nunca baja del 26%.
  * El botón *Margen mínimo* no aplica tasas mayores al 33%: avisa que hay que revisar anticipo,
    residual o plazo.
* **Factores por plazo:** TIIE 8.6% + margen = 30% / 29% / 28% / 27% para 12 / 24 / 36 / 48
  meses. Todos cumplen el rate card con fondeo al 21%.
* **Fuente que ya no está en el catálogo** (por ejemplo, al cargar un folio viejo): se usa la
  fuente principal y aparece una alerta. Así el margen nunca se calcula con fondeo en 0%.

## Fondeo de la operación (versión 3)

* **Dónde se elige (el ejecutivo de cuenta):** en el **Cotizador**, renglón "Datos para el
  cálculo", celda dorada **Fuente de fondeo**. Junto a ella aparece la tasa de fondeo que resulta.
  Si se necesita otra tasa en una operación especial, se captura en **Tasa de fondeo manual**
  (sección de seguro); vacía usa la tasa del catálogo.
* **Dónde se administra (dónde se fondea la empresa):** en **Catálogos** (clave de
  administrador), tabla **FUENTES DE FONDEO**. Cada fuente tiene un tipo de tasa:
  * **Tasa fija.** Ejemplo: recursos propios al 20.5%.
  * **TIIE + spread.** Ejemplo: línea bancaria con TIIE + 4.0%. Se recalcula sola cuando cambia
    la TIIE.

  Hay 6 renglones disponibles. Las líneas bancarias que vienen son ejemplos: capture sus bancos y
  spreads reales.
* **Efecto:** la tasa de fondeo descuenta los flujos para el margen de caja, el spread, la tasa
  mínima y el dictamen del rate card. La fuente queda en la hoja Riesgo y en Factores, y se
  guarda en el Historial con cada folio.

## Botones y seguridad (versión 2, al estilo ABC)

**Botones dentro de cada hoja.** No se imprimen y se distinguen por color:
* Morado: cálculo.
* Dorado: documentos para el cliente.
* Gris: navegación y acceso.

| Hoja | Botones |
|---|---|
| Cotizador | Nueva cotización · Guardar · Cargar folio · Propuesta completa PDF · Generar e-Propuesta · Enviar por correo · Ajustar renta deducible · Acceso gerencia |
| Propuesta / Pago Inicial / Venta | Generar e-documento (libro nuevo solo con valores, como el "Generar propuesta" de ABC) · PDF · Regresar |
| Factores | Ocultar hoja · Ajustar renta deducible · Margen mínimo (rate card) · Expediente interno PDF · Regresar |
| Bonos | Mostrar bonos · Ocultar bonos · Ocultar hoja |
| Riesgo, Tabla, Historial, Catálogos | PDF, cargar folio, proteger, regresar |

* **e-Documentos:** la hoja se copia a un libro nuevo, solo con valores, sin fórmulas, botones ni
  nombres internos. Se guarda como `e-Propuesta_<folio>_<cliente>.xlsx`.
* **Ajustar renta deducible:** calcula la renta extraordinaria exacta para que la renta mensual
  sea 100% deducible: $6,000 para autos ($200 diarios) u $8,550 para eléctricos/híbridos
  ($285 diarios). Equivale al "Ajustar renta $6,000" de ABC, pero sin Buscar objetivo.
* **Hojas restringidas:** Factores, Bonos y Riesgo se abren con la **clave de gerencia**.
  Catálogos y Corridas se abren con la **clave de administrador**. Al abrir el libro se vuelven a
  ocultar.
* **Protección:** la estructura del libro y las hojas de cálculo están protegidas con la clave de
  administrador. Solo las celdas de captura se pueden editar.
* **Contraseñas:** las de esta versión pública son de ejemplo (`FP-Admin2026` / `FP-Gerencia2026`).
  La versión entregada al cliente usa contraseñas distintas. Para cambiarlas, edite las constantes
  `CLAVE_ADMIN` y `CLAVE_GERENCIA` del módulo `modCotizador` (Alt+F11) y vuelva a proteger con el
  botón *Proteger libro y hojas*.

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
