# Cotizador de Arrendamiento – SOFOPLUS

Archivo: **`Cotizador_Arrendamiento_SOFOPLUS.xlsm`** (Excel con macros). Reemplaza al archivo
`COTIZACIÓN LEASING 24 36 48`.

## Cómo empezar

1. Descargue el archivo. Si Windows lo bloquea: clic derecho → *Propiedades* → marque **Desbloquear** → Aceptar.
2. Ábralo y pulse **Habilitar contenido** (macros).
3. Aparece la pestaña **"Cotizador Arrendamiento"** en la cinta de opciones con todos los botones.
   Las macros también se pueden ejecutar con **Alt+F8**.
4. Capture solo las **celdas amarillas con texto azul** de la hoja **Cotizador**.

| Atajo | Acción |
|---|---|
| Ctrl+Shift+N | Nueva cotización (limpia y asigna folio) |
| Ctrl+Shift+G | Guardar en Historial |
| Ctrl+Shift+P | Carta en PDF |

## Hojas

| Hoja | Uso |
|---|---|
| Inicio | Guía rápida, accesos y modelo matemático |
| Cotizador | Captura (cliente, equipo, 4 escenarios), resultados, rentabilidad interna y alertas |
| Carta Cotizacion | Carta para el cliente lista para imprimir/PDF (solo escenarios con "Incluir = Sí") |
| Tabla de Pagos | Calendario de pagos con IVA del escenario elegido |
| Sensibilidad | Matriz de rentas: plazo (12 a 84 meses) × tasa |
| Historial | Registro de cotizaciones; permite recargar cualquier folio |
| Corrida 1 a 4 | Corrida financiera completa (amortización, flujos, VP, TIR, CAT); protegidas |
| Configuracion | Empresa, IVA, fondeo, CM mínimo, promotores, folios, textos legales de la carta |

## Macros

| Macro | Qué hace |
|---|---|
| NuevaCotizacion | Limpia la captura con los valores por defecto de *Configuracion* y el siguiente folio |
| GuardarCotizacion | Guarda/actualiza el folio en *Historial* con un respaldo completo de la captura |
| CargarCotizacion | Recupera un folio del Historial (fila seleccionada o por número) |
| ExportarCartaPDF | Crea `Cotizacion_<folio>_<cliente>.pdf`; avisa si algún escenario no cumple el CM |
| ExportarTablaPDF | Crea el PDF de la tabla de pagos del escenario elegido |
| EnviarPorCorreo | Genera el PDF y abre un correo de Outlook al cliente con el PDF adjunto |
| AplicarTasaMinima | Pone la tasa mínima que cumple el cash margin objetivo (reemplaza a `AjustarCash`) |
| TasaParaRentaDeseada | El cliente pide una renta: calcula y aplica la tasa que la produce |
| CopiarEscenario1 | Copia condiciones del escenario 1 a los escenarios 2–4 |
| ProtegerHojas / DesprotegerHojas | Protección sin contraseña contra borrados accidentales |

El código está en `macros/` por si alguna vez hay que importarlo manualmente
(Alt+F11 → Archivo → Importar archivo).

## Modelo matemático

Notación: *V* valor del equipo sin IVA, *i* = tasa anual / 12, *n* plazo básico, *m* plazo sucesivo,
*tipo* = 1 anticipado / 0 vencido, *f* = fondeo anual / 12.

* **Monto a financiar:** M = V − enganche + seguro financiado + GPS financiado
* **Renta básica:** R = PMT(i, n, −M, VR, tipo), con VR = % residual × V
* **Renta sucesiva:** Rs = PMT(i₂, m, −PF, 0, tipo); si m = 0 el pago final se cobra como opción de compra
* **Amortización:** interés = (saldo − tipo × pago) × i. El saldo al terminar el plazo básico es
  exactamente VR y la tabla termina en 0 (columna de cuadre en cada corrida).
* **Cash margin (C/R):** flujo en firma + Σ flujoₜ × (1 + fondeo/360)^(−días) × (1 + f)^(−t).
  Flujo en firma = −(V + financiados) + enganche + comisión + renta proporcional + depósito − comisión banco.
  El depósito se devuelve en el último pago. **S/R** excluye el pago final/residual.
* **Tasa mínima para el CM objetivo:** el CM es lineal en la renta, CM(R) = A + B·R, así que
  R_mín = (CM objetivo × M − A) / B y la tasa = RATE(n, R_mín, −M, VR, tipo) × 12. Es una solución
  exacta; no requiere *Buscar objetivo*.
* **TIR:** XIRR de los flujos del arrendador con fechas reales. **CAT** informativo: XIRR de los
  flujos del cliente sin IVA (no sustituye el cálculo oficial de Banxico).

## Hallazgos en el archivo original (ya corregidos)

1. **La renta y el pago final no eran consistentes.** La renta se calculaba con un residual de 5 %
   (`Residual (FV)`, la misma celda B20 que el enganche), pero el plazo sucesivo cobraba 40 %
   (`% Residual`). La corrida no cuadraba y la TIR mostrada (52 %) y el cash margin (24 %) estaban inflados.
   Ahora hay un solo **valor residual**; si a propósito se quiere cobrar otro pago final, existe
   el campo opcional *"Pago final distinto al residual"*, que muestra una alerta.
2. **Anticipado/vencido mezclados:** la renta usaba PMT vencido, pero el cash margin cobraba la
   primera renta al inicio. Ahora la modalidad cambia ambos de forma consistente.
3. **La renta proporcional siempre era 0** (la fórmula restaba la misma fecha). Ahora se calcula
   con los días entre la firma y el inicio del plazo.
4. **Cash margin S/R % de 36 y 48 meses** usaba el nombre `cma`, que apunta a la hoja de 24 meses.
5. **Errores #REF!** en la columna de 60 meses de la carta y en la TIR de 36/48 meses.
6. **Carta:** el total de pago inicial de 36 m sumaba el enganche y los de 24/48 m no; GPS de 36 m sin
   IVA aunque decía "IVA incluido".
7. **La renta sucesiva** se calculaba sobre el 40 % del monto financiado en lugar del valor del equipo.

### Importante para la política de precios

Con los datos de ejemplo (tasa 17 %, fondeo 20.5 %) el spread es **negativo**. Con la corrida
consistente el cash margin es negativo en los 4 plazos. La tasa mínima para un cash margin del 8 %
(pago anticipado, enganche 5 %, residual 5 %) es:

| Plazo | 24 m | 36 m | 48 m | 60 m |
|---|---|---|---|---|
| Tasa mínima para CM 8 % | 25.99 % | 24.29 % | 23.45 % | 22.95 % |

Para reproducir la renta exacta del archivo anterior (34,286.77 a 24 meses), use la modalidad
**Vencido**.

## Verificación realizada

* 8,383 fórmulas recalculadas sin ningún error.
* Renta, cash margin C/R y S/R, TIR, CAT, saldo final y tasa mínima comparados contra un modelo
  independiente en Python: coinciden al centavo en los 4 escenarios, y también en casos extremos
  (vencido, opción de compra, seguro financiado, pago final distinto, tasa 0 %, 84 + 12 meses,
  IVA sobre intereses).
* Macros: estructura validada con dos lectores independientes (oletools y LibreOffice) y
  ejecutadas de punta a punta en LibreOffice (guardar, limpiar, restaurar y aplicar tasa mínima).
  El archivo **no se probó en Microsoft Excel**. Si una macro marcara error, importe los módulos de
  `macros/`.

## Regenerar el archivo (opcional, técnico)

La carpeta `fuente/` contiene los scripts que construyen el libro (`build.py`), el proyecto VBA
(`vbabuild.py`) y el ensamblado final (`assemble.py`). Requiere Python 3 con `openpyxl`,
`olefile` y `oletools`, más LibreOffice para el recálculo.
