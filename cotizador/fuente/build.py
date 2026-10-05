import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.drawing.image import Image as XLImage
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
import openpyxl.worksheet.page as page

OUT = 'Cotizador_Arrendamiento.xlsx'
LOGO = 'logo.jpeg'

# ------------------------------------------------------------------ estilo
MORADO = '4B2A7B'
MORADO_CL = 'E9E2F3'
GRIS = 'F2F2F2'
AMARILLO = 'FFF8D6'
FONT = 'Arial'

f_base = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_input = Font(name=FONT, size=10, color='0000FF')
f_link = Font(name=FONT, size=10, color='008000')
f_hdr = Font(name=FONT, size=11, bold=True, color='FFFFFF')
f_title = Font(name=FONT, size=18, bold=True, color=MORADO)
f_sub = Font(name=FONT, size=10, italic=True, color='666666')
f_note = Font(name=FONT, size=9, italic=True, color='666666')
fill_hdr = PatternFill('solid', fgColor=MORADO)
fill_sub = PatternFill('solid', fgColor=MORADO_CL)
fill_in = PatternFill('solid', fgColor=AMARILLO)
fill_gris = PatternFill('solid', fgColor=GRIS)
thin = Side(style='thin', color='BFBFBF')
med = Side(style='medium', color=MORADO)
box = Border(left=thin, right=thin, top=thin, bottom=thin)
topline = Border(top=med)
wrap = Alignment(wrap_text=True, vertical='top')
center = Alignment(horizontal='center', vertical='center', wrap_text=True)
right = Alignment(horizontal='right', vertical='center')
left_c = Alignment(horizontal='left', vertical='center', wrap_text=True)

MON = '$#,##0.00;[Red]-$#,##0.00;"-"'
MON0 = '$#,##0;[Red]-$#,##0;"-"'
PCT = '0.00%;[Red]-0.00%;"-"'
PCT4 = '0.0000%'
FECHA = 'dd/mm/yyyy'
NUM = '#,##0.00'

wb = Workbook()
wb.remove(wb.active)
names = {}


def name(nm, ref):
    names[nm] = ref
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def absref(sheet, rng):
    parts = rng.split(':')
    def a(c):
        col = ''.join(ch for ch in c if ch.isalpha())
        row = ''.join(ch for ch in c if ch.isdigit())
        return '$%s$%s' % (col, row)
    q = "'%s'" % sheet if ' ' in sheet else sheet
    return q + '!' + ':'.join(a(p) for p in parts)


def style_range(ws, rng, font=None, fill=None, border=None, align=None, fmt=None):
    for row in ws[rng]:
        for c in row:
            if font: c.font = font
            if fill: c.fill = fill
            if border: c.border = border
            if align: c.alignment = align
            if fmt: c.number_format = fmt


def header_bar(ws, row, c1, c2, text):
    ws.cell(row=row, column=c1, value=text)
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill_hdr
        cell.font = f_hdr
        cell.alignment = Alignment(vertical='center')
    ws.row_dimensions[row].height = 20


def inp(cell, value, fmt=None, note=None):
    cell.value = value
    cell.font = f_input
    cell.fill = fill_in
    cell.border = box
    cell.protection = Protection(locked=False)
    if fmt:
        cell.number_format = fmt
    if note:
        cell.comment = Comment(note, 'Cotizador')


def calc(cell, formula, fmt=None, bold=False):
    cell.value = formula
    cell.font = f_bold if bold else f_base
    cell.border = box
    if fmt:
        cell.number_format = fmt


def logo(ws, anchor, h=60):
    img = XLImage(LOGO)
    ratio = img.width / img.height
    img.height = h
    img.width = int(h * ratio)
    ws.add_image(img, anchor)


def base_sheet(title, code, tab=None):
    ws = wb.create_sheet(title)
    ws.sheet_properties.codeName = code
    if tab:
        ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    return ws


def dv_list(ws, formula, cells, prompt=None):
    dv = DataValidation(type='list', formula1=formula, allow_blank=True)
    dv.error = 'Seleccione un valor de la lista.'
    dv.errorTitle = 'Valor no válido'
    if prompt:
        dv.prompt = prompt
        dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(cells)


def dv_num(ws, cells, lo, hi, kind='decimal', msg=''):
    dv = DataValidation(type=kind, operator='between', formula1=str(lo), formula2=str(hi), allow_blank=True)
    dv.error = msg or ('Capture un valor entre %s y %s.' % (lo, hi))
    dv.errorTitle = 'Valor fuera de rango'
    ws.add_data_validation(dv)
    dv.add(cells)


# ==================================================================== HOJAS
ws_ini = base_sheet('Inicio', 'Hoja1', MORADO)
ws_cot = base_sheet('Cotizador', 'Hoja2', 'FFC000')
ws_car = base_sheet('Carta Cotizacion', 'Hoja3', '70AD47')
ws_tab = base_sheet('Tabla de Pagos', 'Hoja4', '70AD47')
ws_sen = base_sheet('Sensibilidad', 'Hoja5', '5B9BD5')
ws_his = base_sheet('Historial', 'Hoja6', '5B9BD5')
ws_cr = [base_sheet('Corrida %d' % k, 'Hoja%d' % (6 + k), '808080') for k in range(1, 5)]
ws_cfg = base_sheet('Configuracion', 'Hoja11', '808080')

# ==================================================================== CONFIGURACION
ws = ws_cfg
for col, w in zip('ABCDEFGHIJKLMNOP', [2, 46, 22, 14, 14, 14, 3, 30, 24, 20, 14, 10, 14, 16, 18, 3]):
    ws.column_dimensions[col].width = w
ws['B1'] = 'CONFIGURACIÓN DEL COTIZADOR'
ws['B1'].font = f_title
ws['B2'] = 'Solo el administrador debe modificar esta hoja. Las celdas amarillas son editables.'
ws['B2'].font = f_sub
header_bar(ws, 3, 2, 6, 'DATOS DE LA EMPRESA Y PARÁMETROS')
cfg = [
    ('Razón social del arrendador', 'SOFOPLUS, S.A.P.I. DE C.V., E.R.', 'par_Arrendador', None),
    ('Nombre comercial', 'SOFOPLUS', 'par_Comercial', None),
    ('Domicilio (línea 1)', 'Paseo de los Tamarindos No. 90 Piso 24 Torre 1', 'par_Dir1', None),
    ('Domicilio (línea 2)', 'Col. Bosques de las Lomas', 'par_Dir2', None),
    ('Domicilio (línea 3)', 'CDMX, C.P. 05120', 'par_Dir3', None),
    ('Ciudad para la fecha de la carta', 'Ciudad de México', 'par_Ciudad', None),
    ('Teléfono / correo / web (pie de carta)', '', 'par_Contacto', None),
    ('Tasa de IVA', 0.16, 'par_IVA', PCT),
    ('Tasa de fondeo anual (valor por defecto)', 0.205, 'par_Fondeo', PCT),
    ('Cash margin mínimo objetivo (% del monto financiado)', 0.08, 'par_CMmin', PCT),
    ('Vigencia de la propuesta (días hábiles)', 5, 'par_Vigencia', '0'),
    ('Base de días para costo de fondeo diario', 360, 'par_BaseDias', '0'),
    ('IVA en arrendamiento financiero', 'Sobre renta', 'par_BaseIVAFin', None),
    ('Prefijo del folio', 'COT-', 'par_Prefijo', None),
    ('Último consecutivo utilizado (lo actualiza la macro)', 0, 'par_Consecutivo', '0'),
    ('Carpeta para PDF (vacío = carpeta del archivo)', '', 'par_CarpetaPDF', None),
    ('Correo con copia (CC) al enviar cotizaciones', '', 'par_CorreoCC', None),
]
r = 4
for label, val, nm, fmt in cfg:
    ws.cell(row=r, column=2, value=label).font = f_base
    inp(ws.cell(row=r, column=3), val, fmt)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    name(nm, absref(ws.title, 'C%d' % r))
    r += 1
ws['C4'].comment = Comment('Fuente: carta de cotización del archivo original.', 'Cotizador')
ws['C12'].comment = Comment('Dato del archivo original (CM 24m!B17 = 20.5%). Actualícelo con su costo de fondeo vigente.', 'Cotizador')
ws['C13'].comment = Comment('El archivo original usaba 8% como meta por defecto en la macro AjustarCash.', 'Cotizador')
dv_list(ws, '=lst_BaseIVA', 'C16')

r += 1
DEF_ROW = r + 1
header_bar(ws, r, 2, 6, 'VALORES POR DEFECTO PARA UNA NUEVA COTIZACIÓN')
for k in range(4):
    c = ws.cell(row=r, column=3 + k, value='Escenario %d' % (k + 1))
    c.alignment = center
r += 1
defaults = [
    ('Plazo básico (meses)', [24, 36, 48, 60], 'def_Plazo', '0'),
    ('Plazo sucesivo (meses, 0 = opción de compra)', [1, 1, 1, 1], 'def_Sucesivo', '0'),
    ('Tasa anual plazo básico', [0.17] * 4, 'def_Tasa', PCT),
    ('Tasa anual plazo sucesivo', [0.17] * 4, 'def_TasaSuc', PCT),
    ('Enganche (% del valor)', [0.05] * 4, 'def_Enganche', PCT),
    ('Valor residual (% del valor)', [0.05] * 4, 'def_Residual', PCT),
    ('Comisión por apertura (% del valor)', [0.008] * 4, 'def_Comision', PCT),
    ('Depósito en garantía (# rentas con IVA)', [1] * 4, 'def_Deposito', '0.00'),
]
for label, vals, nm, fmt in defaults:
    ws.cell(row=r, column=2, value=label).font = f_base
    for k, v in enumerate(vals):
        inp(ws.cell(row=r, column=3 + k), v, fmt)
    name(nm, absref(ws.title, 'C%d:F%d' % (r, r)))
    r += 1
singles = [
    ('Tipo de arrendamiento', 'Arrendamiento Puro', 'def_Tipo', '=lst_Tipo'),
    ('Moneda', 'Moneda Nacional', 'def_Moneda', '=lst_Moneda'),
    ('Modalidad de pago', 'Anticipado', 'def_Modalidad', '=lst_Modalidad'),
]
for label, v, nm, lst in singles:
    ws.cell(row=r, column=2, value=label).font = f_base
    inp(ws.cell(row=r, column=3), v)
    name(nm, absref(ws.title, 'C%d' % r))
    dv_list(ws, lst, 'C%d' % r)
    r += 1

# ---- textos de la carta
r += 1
header_bar(ws, r, 2, 6, 'TEXTOS DE LA CARTA (use {ARRENDADOR} y {VIGENCIA} como comodines)')
r += 1
TEXTOS = [
    ('Gastos de Ratificación:', 'Dependiendo de la ciudad, pagaderos (MXN) por cada Anexo de Arrendamiento.'),
    ('Opciones al Término del Contrato:', 'Al concluir el Plazo Básico, o cualquier extensión del mismo, de cada Anexo de Arrendamiento, y siempre y cuando el Arrendatario se encuentre al corriente en el cumplimiento de todas y cada una de las obligaciones que del Arrendamiento Maestro, así como cualesquiera Anexo de Arrendamiento le imponen, este último tendrá derecho preferente para adquirir la totalidad (pero no menos que la totalidad) del Equipo materia del arrendamiento instrumentado en el Anexo de arrendamiento de que se trate. Si el Arrendatario no ejerce su derecho de preferencia mencionado en el párrafo anterior, el Arrendatario deberá, a su costo, desinstalar, re empaquetar y devolver el equipo (no menos que todo) al Arrendador en su domicilio, o el lugar y fecha que este último le indique por escrito para tal efecto, en el estado y condiciones en que el Arrendatario lo hubiese recibido, sin mayor desgaste que el normalmente derivado de su uso o deterioro normal. En caso de que el Equipo no fuese devuelto en la fecha a que se refiere el párrafo anterior, el plazo del arrendamiento respectivo se prorrogará automáticamente trimestralmente hasta la devolución respectiva, en la inteligencia de que las prórrogas correspondientes se considerarán forzosas para efectos del arrendamiento y durante la vigencia de cada una de ellas el Arrendatario se obliga a pagar al Arrendador una cantidad igual al Pago Periódico correspondiente. La totalidad de impuestos, derechos, contribuciones o gastos de cualquier naturaleza que se ocasionen con motivo de la adquisición del Equipo, serán a cargo del Arrendatario.'),
    ('1.- Compra del Equipo:', 'Anticipamos que el Arrendatario enviará al proveedor del equipo su orden de compra. El Arrendador asumirá el pago de la orden de compra del Arrendatario, lo anterior condicionado a que el Arrendatario rente el equipo del Arrendador.'),
    ('2.- Beneficios Fiscales:', 'El Arrendador deberá ser considerado como el dueño del equipo para los efectos fiscales correspondientes.'),
    ('3.- Arrendamiento Neto:', 'El Arrendatario conviene en que su obligación de pagar la Renta, mediante los Pagos Periódicos, y otras cantidades que sean pagaderas conforme al presente, será absoluta e incondicional, y no estará sujeta a reducciones de ninguna clase, incluyendo sin limitación alguna, reducciones por causa de cualquier demanda pasada, presente o futura que sea consecuencia de este Arrendamiento, cualquier Anexo de Arrendamiento o por cualquier otra razón, o en contra del fabricante, proveedor o constructor del Equipo o en contra de cualquier otra persona física o persona moral. El Arrendatario es responsable por todos los gastos de mantenimiento, seguro e impuestos relativos a la compra, renta, posesión y uso del equipo excepto por el Impuesto al Valor Agregado y gastos de importación correspondientes a la compra e importación del Equipo, siendo estos últimos parte del costo del Equipo. El Arrendatario también es responsable del Impuesto al Valor Agregado de las rentas, excluyendo el impuesto sobre la renta a cargo del Arrendador.'),
    ('4.- Moneda:', 'Todos los pagos por concepto de arrendamiento están expresados en la moneda descrita en el tipo de operación así como cualquier otra cantidad que deba ser pagada al Arrendador de acuerdo a la documentación (incluyendo sin limitación alguna costos y gastos).'),
    ('5.- Mantenimiento y Seguro:', 'Todos los mantenimientos y seguros de los equipos serán responsabilidad del Arrendatario durante todo el tiempo que dure el Plazo Básico, o cualquier extensión del mismo establecido en cualesquiera de los Anexos de arrendamiento. El Arrendatario deberá contratar el seguro por los montos y compañía de seguros aceptable al Arrendador y de acuerdo a disponibilidad de coberturas ofrecidas en México. El Arrendatario deberá enviar al Arrendador copia de las pólizas o certificados de seguro como evidencia de la cobertura.'),
    ('6.- Caducidad:', 'La presente propuesta, sus términos y condiciones expiran {VIGENCIA} días hábiles posteriores a la fecha aquí marcada, si {ARRENDADOR} no recibe aceptación por escrito de la misma.'),
    ('7.- Cambio Material Adverso:', 'La presente propuesta está sujeta a que no suceda ningún Cambio Material Adverso que pueda afectar negativamente los términos y condiciones presentados, entendiéndose por cambio material adverso cualquier alteración en las condiciones macroeconómicas, microeconómicas y sociopolíticas del País, liberando por tanto el Arrendatario a el Arrendador de cualquier responsabilidad a este efecto.'),
    ('8.- Confidencialidad:', 'La presente propuesta solo podrá ser vista y analizada por el personal de la Arrendataria y aquellas personas que formen parte del proceso de autorización de esta propuesta.'),
    ('9.- Riesgo Cambiario:', '{ARRENDADOR} no asume ninguna responsabilidad derivada por cualquier variación al tipo de cambio al momento del pago al proveedor, obligándose el Arrendatario a pagar al Arrendador previo pago al proveedor, cualquier diferencia que pudiera existir con relación al tipo de cambio estimado para el cierre de la operación de arrendamiento.'),
    ('10.- Renta Proporcional:', 'El Arrendatario se obliga a pagar una contraprestación por concepto de renta proporcional, misma que se genere por los días que transcurran entre la fecha de pago al proveedor y la fecha de inicio del plazo del arrendamiento.'),
    ('11.- Autorización:', 'La autorización de la presente propuesta está sujeta a que el departamento de crédito y legal de {ARRENDADOR} la aprueben mediante los procesos internos establecidos y la celebración de la documentación correspondiente, por lo que ninguna de las partes involucradas estará legalmente comprometida con la otra por causa de esta propuesta, ni se generará ningún derecho, responsabilidad u obligación como resultado de la misma.'),
    ('Cierre:', 'Agradecemos de antemano el favorecernos con su decisión, esperando tener noticias de usted a la brevedad posible.'),
]
TXT_ROW0 = r
for label, txt in TEXTOS:
    ws.cell(row=r, column=2, value=label).font = f_bold
    ws.cell(row=r, column=2).alignment = wrap
    inp(ws.cell(row=r, column=3), txt)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    ws.cell(row=r, column=3).alignment = wrap
    ws.row_dimensions[r].height = max(15, 12.5 * (len(txt) // 72 + 1))
    r += 1
ws.cell(row=TXT_ROW0 - 1, column=2).comment = Comment('Textos tomados de la hoja "Carta Cotizacion" del archivo original.', 'Cotizador')

# ---- listas
header_bar(ws, 3, 8, 15, 'LISTAS DESPLEGABLES')
listas = [
    ('H', 'Promotores', ['NOMBRE DEL PROMOTOR'] + [''] * 19, 'lst_Promotores', 20),
    ('I', 'Tipo de arrendamiento', ['Arrendamiento Puro', 'Arrendamiento Financiero'], 'lst_Tipo', 2),
    ('J', 'Moneda', ['Moneda Nacional', 'Dólares Americanos'], 'lst_Moneda', 2),
    ('K', 'Modalidad', ['Anticipado', 'Vencido'], 'lst_Modalidad', 2),
    ('L', 'Sí / No', ['Sí', 'No'], 'lst_SiNo', 2),
    ('M', 'Forma de pago', ['Financiado', 'Contado'], 'lst_Forma', 2),
    ('N', 'Base IVA financiero', ['Sobre renta', 'Sobre intereses'], 'lst_BaseIVA', 2),
    ('O', 'Estatus', ['Enviada', 'En seguimiento', 'Aceptada', 'Rechazada', 'Vencida'], 'lst_Estatus', 5),
]
for col, title, vals, nm, n in listas:
    c = ws['%s4' % col]
    c.value = title
    c.font = f_bold
    c.fill = fill_sub
    for i, v in enumerate(vals):
        cell = ws['%s%d' % (col, 5 + i)]
        if nm == 'lst_Promotores':
            inp(cell, v if v else None)
        else:
            cell.value = v
            cell.font = f_base
            cell.border = box
    name(nm, absref(ws.title, '%s5:%s%d' % (col, col, 4 + n)))
ws['H26'] = 'Agregue o quite promotores en la lista amarilla.'
ws['H26'].font = f_note
ws.freeze_panes = 'A4'

# ==================================================================== COTIZADOR
ws = ws_cot
for col, w in zip('ABCDEFGH', [2, 46, 18, 18, 18, 18, 2, 60]):
    ws.column_dimensions[col].width = w
logo(ws, 'B1', 58)
ws['C1'] = 'COTIZADOR DE ARRENDAMIENTO'
ws['C1'].font = f_title
ws.merge_cells('C1:F1')
ws['C2'] = 'Corridas financieras, cash margin, TIR y carta para el cliente'
ws['C2'].font = f_sub
ws.merge_cells('C2:F2')
ws['C3'] = '="Folio: "&inp_Folio&"   |   "&inp_Cliente'
ws['C3'].font = Font(name=FONT, size=11, bold=True, color=MORADO)
ws.merge_cells('C3:F3')
for rr in (1, 2, 3):
    ws.row_dimensions[rr].height = 22
ws['H1'] = 'CÓMO USAR'
ws['H1'].font = f_bold
ws['H2'] = '1) Capture las celdas AMARILLAS (texto azul). 2) Revise resultados y semáforo. 3) Use la pestaña "Cotizador Arrendamiento" de la cinta (o Alt+F8) para guardar, generar PDF o enviar.'
ws['H2'].alignment = wrap
ws['H2'].font = f_note
ws.merge_cells('H2:H4')

header_bar(ws, 5, 2, 8, '1. DATOS DEL CLIENTE')
hoy = dt.datetime(2026, 10, 5)
datos = [
    ('Folio de la cotización', 'COT-0001', 'inp_Folio', None, 'La macro "Nueva cotización" asigna el siguiente folio automáticamente.'),
    ('Fecha de la cotización', hoy, 'inp_Fecha', FECHA, 'Fecha que aparece en la carta.'),
    ('Cliente / Arrendatario (razón social)', 'CLIENTE DE EJEMPLO, S.A. DE C.V.', 'inp_Cliente', None, 'Razón social completa del arrendatario.'),
    ('RFC del cliente', '', 'inp_RFC', None, 'Opcional.'),
    ('Atención a (nombre del contacto)', '', 'inp_Contacto', None, 'Opcional. Aparece en la carta como "At\'n".'),
    ('Correo electrónico del cliente', '', 'inp_Correo', None, 'Se usa para el botón "Enviar por correo".'),
    ('Proveedor del equipo', '', 'inp_Proveedor', None, ''),
    ('Descripción del equipo', 'Auto BYD M9 2026', 'inp_Equipo', None, 'Marca, modelo, año, cantidad.'),
    ('Promotor', 'NOMBRE DEL PROMOTOR', 'inp_Promotor', None, 'Lista editable en la hoja Configuracion.'),
    ('Obligado solidario', 'Por definir', 'inp_Obligado', None, ''),
]
r = 6
for label, val, nm, fmt, note in datos:
    ws.cell(row=r, column=2, value=label).font = f_base
    inp(ws.cell(row=r, column=3), val, fmt)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    name(nm, absref(ws.title, 'C%d' % r))
    ws.cell(row=r, column=8, value=note).font = f_note
    r += 1
dv_list(ws, '=lst_Promotores', names['inp_Promotor'].split('!')[1].replace('$', ''))

r += 1
header_bar(ws, r, 2, 8, '2. CONDICIONES GENERALES DE LA OPERACIÓN')
r += 1
gen = [
    ('Tipo de arrendamiento', 'Arrendamiento Puro', 'inp_Tipo', None, '=lst_Tipo', 'Puro: IVA sobre la renta completa. Financiero: según Configuracion.'),
    ('Moneda', 'Moneda Nacional', 'inp_Moneda', None, '=lst_Moneda', 'Todos los montos se capturan en la moneda de la operación.'),
    ('Tipo de cambio (informativo)', 1, 'inp_TC', '#,##0.0000', None, 'Solo informativo; capture 1 si es Moneda Nacional.'),
    ('Modalidad de pago de las rentas', 'Anticipado', 'inp_Modalidad', None, '=lst_Modalidad', 'Anticipado: la renta se paga al inicio de cada mes. Vencido: al final. Cambia la fórmula de la renta (PMT tipo 1 ó 0).'),
    ('Valor del equipo (sin IVA)', 758450, 'inp_Valor', MON, None, 'Precio de factura del proveedor antes de IVA (ejemplo: archivo original).'),
    ('Fecha de firma / pago al proveedor', dt.datetime(2026, 9, 30), 'inp_FechaFirma', FECHA, None, 'Fecha del desembolso; aquí se cobran enganche, comisión, depósito y renta proporcional.'),
    ('Fecha de la primera renta', dt.datetime(2026, 10, 1), 'inp_FechaPrimera', FECHA, None, 'Si hay días entre la firma y el inicio del plazo se cobra renta proporcional (renta/30 x días).'),
    ('Comisión banco / alianza ($, uso interno)', 0, 'inp_ComBanco', MON, None, 'Costo pagado a un tercero; reduce el cash margin.'),
    ('Tasa de fondeo anual (uso interno)', 0.205, 'inp_Fondeo', PCT, None, 'Costo del dinero para calcular cash margin. Valor inicial desde Configuracion.'),
    ('Cash margin objetivo (% del monto financiado)', 0.08, 'inp_CMobj', PCT, None, 'Meta interna; con ella se calcula la tasa mínima por escenario.'),
]
for label, val, nm, fmt, lst, note in gen:
    ws.cell(row=r, column=2, value=label).font = f_base
    inp(ws.cell(row=r, column=3), val, fmt)
    name(nm, absref(ws.title, 'C%d' % r))
    if lst:
        dv_list(ws, lst, 'C%d' % r)
    ws.cell(row=r, column=8, value=note).font = f_note
    ws.cell(row=r, column=8).alignment = Alignment(wrap_text=True, vertical='center')
    r += 1
    if nm == 'inp_Valor':
        ws.cell(row=r, column=2, value='   IVA del equipo').font = f_base
        calc(ws.cell(row=r, column=3), '=inp_Valor*par_IVA', MON)
        r += 1
        ws.cell(row=r, column=2, value='   Valor del equipo con IVA').font = f_base
        calc(ws.cell(row=r, column=3), '=inp_Valor+C%d' % (r - 1), MON)
        ws.cell(row=r, column=4, value='=IF(inp_TC<>1,"Equivalente MXN: "&TEXT(C%d*inp_TC,"$#,##0.00"),"")' % r).font = f_note
        r += 1
dv_num(ws, names['inp_Valor'].split('!')[1].replace('$', ''), 0, 1e12)
dv_num(ws, names['inp_Fondeo'].split('!')[1].replace('$', ''), 0, 1)
dv_num(ws, names['inp_CMobj'].split('!')[1].replace('$', ''), -1, 1)

r += 1
header_bar(ws, r, 2, 8, '3. ESCENARIOS A COTIZAR (hasta 4 plazos en la misma carta)')
r += 1
ESC_HDR = r
for k in range(4):
    c = ws.cell(row=r, column=3 + k, value='Escenario %d' % (k + 1))
    c.font = f_bold
    c.fill = fill_sub
    c.alignment = center
    c.border = box
ws.cell(row=r, column=2, value='Concepto').font = f_bold
ws.cell(row=r, column=2).fill = fill_sub
r += 1
esc = [
    ('Incluir en la carta', ['Sí'] * 4, 'esc_Incluir', None, '=lst_SiNo', 'Elija "No" para ocultar el escenario en la carta.'),
    ('Plazo básico (meses)', [24, 36, 48, 60], 'esc_Plazo', '0', ('whole', 1, 84), '1 a 84 meses.'),
    ('Plazo sucesivo (meses; 0 = opción de compra)', [1] * 4, 'esc_Sucesivo', '0', ('whole', 0, 12), 'El valor residual se cobra en rentas del plazo sucesivo. Con 0 se cobra como opción de compra única.'),
    ('Tasa anual plazo básico', [0.17] * 4, 'esc_Tasa', PCT, ('decimal', 0, 2), 'Tasa nominal anual; se divide entre 12 (pagos mensuales).'),
    ('Tasa anual plazo sucesivo', [0.17] * 4, 'esc_TasaSuc', PCT, ('decimal', 0, 2), ''),
    ('Enganche / downpayment (% del valor)', [0.05] * 4, 'esc_Enganche', PCT, ('decimal', 0, 0.9), 'Archivo original: 5%.'),
    ('Valor residual (% del valor)', [0.05] * 4, 'esc_Residual', PCT, ('decimal', 0, 0.9), 'Saldo que queda al final del plazo básico (valor futuro en la fórmula de renta).'),
    ('Pago final distinto al residual (% del valor, opcional)', [None] * 4, 'esc_PagoFinal', PCT, ('decimal', 0, 0.9), 'Déjelo VACÍO para que el pago final sea igual al residual (corrida que cuadra en cero).'),
    ('Comisión por apertura (% del valor)', [0.008] * 4, 'esc_Comision', PCT, ('decimal', 0, 0.2), 'Archivo original: 0.8% sobre el valor del equipo.'),
    ('Depósito en garantía (# de rentas con IVA)', [1] * 4, 'esc_Deposito', '0.00', ('decimal', 0, 12), 'Se devuelve/aplica en el último pago.'),
    ('Seguro del equipo (monto total sin IVA)', [0] * 4, 'esc_Seguro', MON, ('decimal', 0, 1e12), 'Monto de la póliza por todo el plazo.'),
    ('Seguro: forma de pago', ['Financiado'] * 4, 'esc_SeguroForma', None, '=lst_Forma', 'Financiado: se suma al monto a financiar. Contado: se cobra en el pago inicial.'),
    ('GPS (monto total sin IVA)', [0] * 4, 'esc_GPS', MON, ('decimal', 0, 1e12), ''),
    ('GPS: forma de pago', ['Financiado'] * 4, 'esc_GPSForma', None, '=lst_Forma', ''),
]
ESC = {}
for label, vals, nm, fmt, val, note in esc:
    ws.cell(row=r, column=2, value=label).font = f_base
    for k, v in enumerate(vals):
        inp(ws.cell(row=r, column=3 + k), v, fmt)
        ws.cell(row=r, column=3 + k).alignment = Alignment(horizontal='center')
    name(nm, absref(ws.title, 'C%d:F%d' % (r, r)))
    ESC[nm] = r
    if isinstance(val, str):
        dv_list(ws, val, 'C%d:F%d' % (r, r))
    elif val:
        dv_num(ws, 'C%d:F%d' % (r, r), val[1], val[2], val[0])
    ws.cell(row=r, column=8, value=note).font = f_note
    ws.cell(row=r, column=8).alignment = Alignment(wrap_text=True, vertical='center')
    r += 1

# ---- resultados (se llenan después de definir celdas de las corridas)
r += 1
RES_CLIENTE = r
header_bar(ws, r, 2, 8, '4. RESULTADOS PARA EL CLIENTE')
r += 1

# ==================================================================== CORRIDAS
FIRST = 46          # primera fila de pagos
NROWS = 98          # filas de pagos (84 + 12 + holgura)
LAST = FIRST + NROWS - 1
T0 = FIRST - 1

P = {}  # etiquetas -> celda de parámetro (columna C)


def build_corrida(ws, k):
    for col, w in zip('ABCDEFGHIJKLMNOPQRS', [7, 17, 7, 12, 16, 15, 14, 15, 16, 13, 16, 16, 11, 16, 16, 12, 16, 16, 16]):
        ws.column_dimensions[col].width = w
    ws.column_dimensions['B'].width = 44
    ws['A1'] = '="CORRIDA FINANCIERA – ESCENARIO "&$C$3&IF($C$4="Sí",""," (no incluido en la carta)")'
    ws['A1'].font = f_title
    ws['A2'] = 'Hoja de cálculo interna (protegida). Todos los datos vienen de la hoja Cotizador.'
    ws['A2'].font = f_sub
    params = [
        ('Escenario', k, None, 'esc'),
        ('Incluido en la carta', '=INDEX(esc_Incluir,1,$C$3)', None, 'incl'),
        ('Plazo básico (n, meses)', '=INDEX(esc_Plazo,1,$C$3)', '0', 'n'),
        ('Plazo sucesivo (m, meses)', '=INDEX(esc_Sucesivo,1,$C$3)', '0', 'm'),
        ('Tipo de pago (1 = anticipado, 0 = vencido)', '=IF(inp_Modalidad="Anticipado",1,0)', '0', 'tau'),
        ('Tasa anual plazo básico', '=INDEX(esc_Tasa,1,$C$3)', PCT4, 'ia'),
        ('Tasa mensual plazo básico (i)', '=C8/12', PCT4, 'i'),
        ('Tasa anual plazo sucesivo', '=INDEX(esc_TasaSuc,1,$C$3)', PCT4, 'i2a'),
        ('Tasa mensual plazo sucesivo (i2)', '=C10/12', PCT4, 'i2'),
        ('Tasa de fondeo anual', '=inp_Fondeo', PCT4, 'fa'),
        ('Tasa de fondeo mensual (f)', '=C12/12', PCT4, 'f'),
        ('Tasa de IVA', '=par_IVA', PCT, 'iva'),
        ('Valor del equipo sin IVA (V)', '=inp_Valor', MON, 'V'),
        ('Enganche', '=C15*INDEX(esc_Enganche,1,$C$3)', MON, 'eng'),
        ('Seguro financiado', '=IF(INDEX(esc_SeguroForma,1,$C$3)="Financiado",INDEX(esc_Seguro,1,$C$3),0)', MON, 'segf'),
        ('GPS financiado', '=IF(INDEX(esc_GPSForma,1,$C$3)="Financiado",INDEX(esc_GPS,1,$C$3),0)', MON, 'gpsf'),
        ('Monto a financiar (M = V − enganche + financiados)', '=C15-C16+C17+C18', MON, 'M'),
        ('Valor residual (% del valor)', '=INDEX(esc_Residual,1,$C$3)', PCT, 'resp'),
        ('Valor residual (VR)', '=C15*C20', MON, 'VR'),
        ('Pago final (% del valor)', '=IF(INDEX(esc_PagoFinal,1,$C$3)="",C20,INDEX(esc_PagoFinal,1,$C$3))', PCT, 'pfp'),
        ('Pago final ($)', '=C15*C22', MON, 'PF'),
        ('RENTA BÁSICA  R = PMT(i, n, −M, VR, tipo)', '=PMT(C9,C5,-C19,C21,C7)', MON, 'R'),
        ('Renta plazo sucesivo  Rs = PMT(i2, m, −PF, 0, tipo)', '=IF(C6>0,PMT(C11,C6,-C23,0,C7),0)', MON, 'Rs'),
        ('Opción de compra (si m = 0)', '=IF(C6=0,C23,0)', MON, 'OC'),
        ('Comisión por apertura (%)', '=INDEX(esc_Comision,1,$C$3)', PCT, 'comp'),
        ('Comisión por apertura ($, sobre el valor)', '=C15*C27', MON, 'com'),
        ('Depósito en garantía (# rentas con IVA)', '=INDEX(esc_Deposito,1,$C$3)', '0.00', 'depn'),
        ('Depósito en garantía ($)', '=C29*C24*(1+C14)', MON, 'dep'),
        ('Fecha de firma / desembolso', '=inp_FechaFirma', FECHA, 'ff'),
        ('Fecha de la primera renta', '=inp_FechaPrimera', FECHA, 'fp'),
        ('Inicio del plazo (t = 0)', '=IF(C7=1,C32,EDATE(C32,-1))', FECHA, 'fi'),
        ('Días entre firma e inicio del plazo', '=MAX(0,C33-C31)', '0', 'd'),
        ('Renta proporcional (R/30 × días)', '=C24/30*C34', MON, 'rp'),
        ('Comisión banco / alianza', '=inp_ComBanco', MON, 'cb'),
        ('Factor de fondeo por días  (1+fa/base)^(−días)', '=(1+C12/par_BaseDias)^(-C34)', '0.000000', 'DF'),
        ('Número de pagos en la tabla', '=C5+IF(C6>0,C6,IF(C23>0,1,0))', '0', 'N'),
        ('Desembolso del arrendador en la firma', '=-(C15+C17+C18)', MON, 'out0'),
        ('Cobros en la firma (enganche+comisión+renta prop.+depósito−com. banco)', '=C16+C28+C35+C30-C36', MON, 'in0'),
        ('Flujo neto del arrendador en la firma', '=C39+C40', MON, 'net0'),
    ]
    rr = 3
    for label, f, fmt, key in params:
        ws.cell(row=rr, column=2, value=label).font = f_base
        c = ws.cell(row=rr, column=3)
        calc(c, f, fmt, bold=key in ('R', 'M'))
        if key == 'esc':
            c.font = f_input
        P[key] = 'C%d' % rr
        rr += 1
    assert P['net0'] == 'C41', P['net0']
    assert P['R'] == 'C24' and P['N'] == 'C38' and P['DF'] == 'C37'

    # ---- resultados
    ws['E3'] = 'RESULTADOS DEL ESCENARIO'
    for c in range(5, 10):
        ws.cell(row=3, column=c).fill = fill_hdr
        ws.cell(row=3, column=c).font = f_hdr
    rows = [
        ('Cash margin C/R ($) = VP al fondeo de todos los flujos', '=C41+SUM(N%d:N%d)' % (FIRST, LAST), MON, 'CM'),
        ('Cash margin C/R (% del monto financiado)', '=IFERROR(I4/C19,0)', PCT, 'CMp'),
        ('Cash margin S/R ($) = sin pago final / residual', '=C41+SUM(S%d:S%d)' % (FIRST, LAST), MON, 'CMs'),
        ('Cash margin S/R (%)', '=IFERROR(I6/C19,0)', PCT, 'CMsp'),
        ('TIR anual efectiva del arrendador (XIRR)', '=IFERROR(XIRR(Q%d:Q%d,P%d:P%d),"n/d")' % (T0, LAST, T0, LAST), PCT, 'TIR'),
        ('TIR nominal anual (capitalizable mensual)', '=IFERROR(12*((1+I8)^(1/12)-1),"n/d")', PCT, 'TIRn'),
        ('CAT informativo sin IVA (flujos del cliente)', '=IFERROR(XIRR(R%d:R%d,P%d:P%d),"n/d")' % (T0, LAST, T0, LAST), PCT, 'CAT'),
        ('Σ factores de descuento de rentas básicas', '=SUMIF(B%d:B%d,"Básico",M%d:M%d)' % (FIRST, LAST, FIRST, LAST), '0.000000', 'SB'),
        ('Factor de descuento del último pago', '=IFERROR(INDEX(M%d:M%d,C38),0)' % (FIRST, LAST), '0.000000', 'ML'),
        ('Sensibilidad del CM a la renta (∂CM/∂R)', '=I11+C29*(1+C14)*(1-I12)+C34/30', '0.000000', 'B'),
        ('Renta mínima para el CM objetivo', '=IFERROR((inp_CMobj*C19-(I4-C24*I13))/I13,0)', MON, 'Rmin'),
        ('Tasa anual mínima para el CM objetivo', '=IFERROR(RATE(C5,I14,-C19,C21,C7)*12,"n/d")', PCT, 'TasaMin'),
        ('Saldo final de la tabla (debe ser 0)', '=IFERROR(INDEX(I%d:I%d,C38),0)' % (FIRST, LAST), MON, 'Saldo'),
        ('Pago inicial: base sin IVA (enganche+comisión+renta prop.+contado)', '=C16+C28+C35+IF(INDEX(esc_SeguroForma,1,$C$3)="Contado",INDEX(esc_Seguro,1,$C$3),0)+IF(INDEX(esc_GPSForma,1,$C$3)="Contado",INDEX(esc_GPS,1,$C$3),0)', MON, 'PIb'),
        ('Pago inicial: IVA', '=I17*C14', MON, 'PIi'),
        ('Pago inicial: depósito en garantía', '=C30', MON, 'PId'),
        ('PAGO INICIAL TOTAL (con IVA)', '=I17+I18+I19', MON, 'PI'),
        ('IVA de la primera renta', '=IFERROR(J%d,0)' % FIRST, MON, 'IVA1'),
        ('Total de pagos con IVA (rentas + pago final)', '=SUM(K%d:K%d)' % (FIRST, LAST), MON, 'TotK'),
        ('Spread (tasa cliente − tasa fondeo)', '=C8-C12', PCT, 'Spread'),
    ]
    rr = 4
    for label, f, fmt, key in rows:
        ws.cell(row=rr, column=5, value=label).font = f_base
        ws.merge_cells(start_row=rr, start_column=5, end_row=rr, end_column=8)
        calc(ws.cell(row=rr, column=9), f, fmt, bold=key in ('CM', 'CMp', 'TIR', 'PI'))
        P[key] = 'I%d' % rr
        rr += 1
    assert P['CM'] == 'I4' and P['CMp'] == 'I5' and P['CMs'] == 'I6' and P['TIR'] == 'I8'
    assert P['SB'] == 'I11' and P['ML'] == 'I12' and P['B'] == 'I13' and P['Rmin'] == 'I14' and P['PIb'] == 'I17'

    ws['E26'] = ('Método: CM = flujo en firma + Σ flujo_t × (1+fa/base)^(−días) × (1+f)^(−t). '
                 'Como el CM es lineal en la renta, la renta mínima se despeja en forma cerrada y la tasa mínima '
                 'se obtiene con RATE. La tabla usa interés = (saldo − tipo × pago) × i, de modo que el saldo al término '
                 'del plazo básico es exactamente el valor residual.')
    ws['E26'].font = f_note
    ws['E26'].alignment = wrap
    ws.merge_cells('E26:I32')

    # ---- tabla
    hdr = ['#', 'Concepto', 't', 'Fecha', 'Saldo inicial', 'Pago s/IVA', 'Interés', 'Capital', 'Saldo final',
           'IVA', 'Total c/IVA', 'Flujo arrendador', 'Factor desc.', 'VP flujo', 'Flujo S/R', 'Fecha (XIRR)',
           'Flujo XIRR', 'Flujo cliente (CAT)', 'VP flujo S/R']
    for j, h in enumerate(hdr):
        c = ws.cell(row=T0 - 1, column=1 + j, value=h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = center
    ws.row_dimensions[T0 - 1].height = 30
    # fila 0 (firma)
    t0 = {1: 0, 2: 'Firma / desembolso', 3: 0, 4: '=C31', 12: '=C41', 16: '=C31', 17: '=C41',
          18: '=C19-C28-C35'}
    for col, v in t0.items():
        c = ws.cell(row=T0, column=col, value=v)
        c.font = f_bold
        c.fill = fill_gris
    for col in range(1, 20):
        ws.cell(row=T0, column=col).fill = fill_gris
    for kk in range(1, NROWS + 1):
        r_ = FIRST + kk - 1
        prev = r_ - 1
        f = {
            'A': kk,
            'B': '=IF(A{r}>$C$38,"",IF(A{r}<=$C$5,"Básico",IF($C$6>0,"Sucesivo","Opción de compra")))',
            'C': '=IF(B{r}="","",IF(B{r}="Opción de compra",$C$5,A{r}-$C$7))',
            'D': '=IF(B{r}="","",EDATE($C$33,C{r}))',
            'E': ('=IF(B{r}="","",$C$19)' if kk == 1 else '=IF(B{r}="","",I{p})'),
            'F': '=IF(B{r}="","",IF(B{r}="Básico",$C$24,IF(B{r}="Sucesivo",$C$25,$C$26)))',
            'G': '=IF(B{r}="","",IF(B{r}="Opción de compra",0,(E{r}-$C$7*F{r})*IF(B{r}="Sucesivo",$C$11,$C$9)))',
            'H': '=IF(B{r}="","",F{r}-G{r})',
            'I': '=IF(B{r}="","",E{r}-H{r})',
            'J': '=IF(B{r}="","",IF(OR(inp_Tipo="Arrendamiento Puro",par_BaseIVAFin="Sobre renta",B{r}="Opción de compra"),F{r},G{r})*$C$14)',
            'K': '=IF(B{r}="","",F{r}+J{r})',
            'L': '=IF(B{r}="","",F{r}-IF(A{r}=$C$38,$C$30,0))',
            'M': '=IF(B{r}="","",$C$37*(1+$C$13)^(-C{r}))',
            'N': '=IF(B{r}="","",L{r}*M{r})',
            'O': '=IF(B{r}="","",IF(B{r}="Básico",F{r},0)-IF(A{r}=$C$38,$C$30,0))',
            'P': '=IF(B{r}="",$C$31,D{r})',
            'Q': '=IF(B{r}="",0,L{r})',
            'R': '=IF(B{r}="",0,-F{r})',
            'S': '=IF(B{r}="","",O{r}*M{r})',
        }
        for col, form in f.items():
            c = ws['%s%d' % (col, r_)]
            c.value = form.format(r=r_, p=prev) if isinstance(form, str) else form
            c.font = f_base
            if col in 'EFGHIJKLNOQRS':
                c.number_format = MON
            elif col in 'DP':
                c.number_format = FECHA
            elif col == 'M':
                c.number_format = '0.000000'
    ws.freeze_panes = 'C%d' % T0
    ws.conditional_formatting.add('A%d:S%d' % (FIRST, LAST),
                                  FormulaRule(formula=['$B%d="Sucesivo"' % FIRST], fill=PatternFill('solid', fgColor='FFF2CC')))
    ws.conditional_formatting.add('A%d:S%d' % (FIRST, LAST),
                                  FormulaRule(formula=['$B%d="Opción de compra"' % FIRST], fill=PatternFill('solid', fgColor='FFF2CC')))
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = '%d:%d' % (T0 - 1, T0 - 1)


for k, w in enumerate(ws_cr, start=1):
    build_corrida(w, k)


def cr(key, k):
    return "'Corrida %d'!%s" % (k, P[key].replace('C', '$C$', 1) if P[key][0] == 'C' else P[key].replace('I', '$I$', 1))


# ==================================================================== COTIZADOR: resultados
ws = ws_cot
for k in range(4):
    c = ws.cell(row=RES_CLIENTE, column=3 + k, value='Escenario %d' % (k + 1))
    c.font = f_hdr
    c.alignment = center
res = [
    ('Plazo', lambda k: '=INDEX(esc_Plazo,1,%d)&" + "&INDEX(esc_Sucesivo,1,%d)&" meses"' % (k, k), None, None, False),
    ('Monto a financiar', lambda k: '=%s' % cr('M', k), MON, 'res_Monto', False),
    ('Renta mensual (sin IVA)', lambda k: '=%s' % cr('R', k), MON, 'res_Renta', True),
    ('IVA de la renta', lambda k: '=%s' % cr('IVA1', k), MON, 'res_IVA', False),
    ('Renta mensual con IVA', lambda k: '=%s+%s' % (cr('R', k), cr('IVA1', k)), MON, 'res_RentaIVA', True),
    ('Renta plazo sucesivo / opción de compra (sin IVA)', lambda k: '=%s+%s' % (cr('Rs', k), cr('OC', k)), MON, 'res_Final', False),
    ('Pago inicial total (con IVA)', lambda k: '=%s' % cr('PI', k), MON, 'res_PagoInicial', True),
    ('Total de pagos del contrato (con IVA)', lambda k: '=%s' % cr('TotK', k), MON, 'res_Total', False),
    ('CAT informativo sin IVA', lambda k: '=%s' % cr('CAT', k), PCT, 'res_CAT', False),
]
r = RES_CLIENTE + 1
for label, fn, fmt, nm, bold in res:
    ws.cell(row=r, column=2, value=label).font = f_bold if bold else f_base
    for k in range(1, 5):
        c = ws.cell(row=r, column=2 + k)
        calc(c, fn(k), fmt, bold)
        c.font = Font(name=FONT, size=10, bold=bold, color='008000')
        c.alignment = Alignment(horizontal='center')
    if nm:
        name(nm, absref(ws.title, 'C%d:F%d' % (r, r)))
    r += 1
ws.cell(row=r - 1, column=8, value='CAT calculado con XIRR sobre los flujos del cliente; informativo, no sustituye el cálculo oficial de Banxico.').font = f_note
ws.cell(row=r - 1, column=8).alignment = Alignment(wrap_text=True, vertical='center')

r += 1
RES_INT = r
header_bar(ws, r, 2, 8, '5. ANÁLISIS INTERNO – RENTABILIDAD (no aparece en la carta)')
for k in range(4):
    c = ws.cell(row=r, column=3 + k, value='Escenario %d' % (k + 1))
    c.alignment = center
r += 1
intr = [
    ('Spread (tasa cliente − fondeo)', lambda k: '=%s' % cr('Spread', k), PCT, None),
    ('TIR anual efectiva del arrendador', lambda k: '=%s' % cr('TIR', k), PCT, 'res_TIR'),
    ('TIR nominal anual (cap. mensual)', lambda k: '=%s' % cr('TIRn', k), PCT, None),
    ('Cash margin C/R ($)', lambda k: '=%s' % cr('CM', k), MON, 'res_CM'),
    ('Cash margin C/R (% del monto financiado)', lambda k: '=%s' % cr('CMp', k), PCT, 'res_CMpct'),
    ('Cash margin S/R ($)', lambda k: '=%s' % cr('CMs', k), MON, None),
    ('Cash margin S/R (%)', lambda k: '=%s' % cr('CMsp', k), PCT, None),
    ('Dictamen vs. CM objetivo', lambda k: '=IF(%s>=inp_CMobj,"CUMPLE","NO CUMPLE")' % cr('CMp', k), None, 'res_Dictamen'),
    ('Renta mínima para el CM objetivo (sin IVA)', lambda k: '=%s' % cr('Rmin', k), MON, None),
    ('Tasa anual mínima para el CM objetivo', lambda k: '=%s' % cr('TasaMin', k), PCT, 'res_TasaMin'),
    ('Cuadre de la corrida (saldo final)', lambda k: '=IF(ABS(%s)<0.01,"OK","Dif. "&TEXT(%s,"#,##0.00"))' % (cr('Saldo', k), cr('Saldo', k)), None, None),
]
for label, fn, fmt, nm in intr:
    ws.cell(row=r, column=2, value=label).font = f_base
    for k in range(1, 5):
        c = ws.cell(row=r, column=2 + k)
        calc(c, fn(k), fmt)
        c.font = Font(name=FONT, size=10, color='008000')
        c.alignment = Alignment(horizontal='center')
    if nm:
        name(nm, absref(ws.title, 'C%d:F%d' % (r, r)))
    if label.startswith('Dictamen'):
        DICT_ROW = r
    r += 1
ws.conditional_formatting.add('C%d:F%d' % (DICT_ROW, DICT_ROW), CellIsRule(operator='equal', formula=['"CUMPLE"'], fill=PatternFill('solid', fgColor='C6EFCE'), font=Font(name=FONT, bold=True, color='006100')))
ws.conditional_formatting.add('C%d:F%d' % (DICT_ROW, DICT_ROW), CellIsRule(operator='equal', formula=['"NO CUMPLE"'], fill=PatternFill('solid', fgColor='FFC7CE'), font=Font(name=FONT, bold=True, color='9C0006')))
ws.cell(row=RES_INT + 1, column=8, value='C/R = con residual (incluye pago final). S/R = sin residual. Descontado a la tasa de fondeo, igual que el archivo original.').font = f_note
ws.cell(row=RES_INT + 1, column=8).alignment = Alignment(wrap_text=True, vertical='center')
ws.cell(row=DICT_ROW + 2, column=8, value='Macro "Aplicar tasa mínima": copia esta tasa a la fila de tasas del escenario.').font = f_note

r += 1
header_bar(ws, r, 2, 8, '6. VALIDACIONES Y ALERTAS')
r += 1
ALERT_ROW = r
for k in range(1, 5):
    E = lambda nm: 'INDEX(%s,1,%d)' % (nm, k)
    f = ('=IF(inp_Valor<=0,"Capture el valor del equipo. ","")'
         '&IF(inp_FechaPrimera<inp_FechaFirma,"La 1a renta es anterior a la firma. ","")'
         '&IF({tasa}<inp_Fondeo,"Spread negativo (tasa < fondeo). ","")'
         '&IF({tasa}=0,"Tasa en cero. ","")'
         '&IF({res}+{eng}>=1,"Enganche + residual >= 100%. ","")'
         '&IF(AND({pf}<>"",{pf}<>{res}),"Pago final distinto al residual: la tasa real difiere de la pactada. ","")'
         '&IF({cmp}<inp_CMobj,"CM por debajo del objetivo. ","")'
         '&IF(ABS({saldo})>=0.01,"La corrida no cuadra. ","")').format(
        tasa=E('esc_Tasa'), res=E('esc_Residual'), eng=E('esc_Enganche'), pf=E('esc_PagoFinal'),
        cmp=cr('CMp', k), saldo=cr('Saldo', k))
    f = '=IF(%s="","Sin alertas",%s)' % (f[1:], f[1:])
    c = ws.cell(row=r, column=2 + k, value=f)
    c.font = Font(name=FONT, size=9, color='C00000')
    c.alignment = wrap
    c.border = box
ws.cell(row=r, column=2, value='Alertas por escenario').font = f_bold
ws.row_dimensions[r].height = 75
ws.conditional_formatting.add('C%d:F%d' % (r, r), CellIsRule(operator='equal', formula=['"Sin alertas"'], font=Font(name=FONT, color='006100')))
name('res_Alertas', absref(ws.title, 'C%d:F%d' % (r, r)))
ws.freeze_panes = 'C5'
ws.print_area = 'A1:H%d' % r
ws.page_setup.orientation = 'portrait'
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True

# ==================================================================== CARTA
ws = ws_car
for col, w in zip('ABCDEFG', [2, 40, 17, 17, 17, 17, 2]):
    ws.column_dimensions[col].width = w
logo(ws, 'B1', 62)
for i, nm in enumerate(['par_Arrendador', 'par_Dir1', 'par_Dir2', 'par_Dir3'], start=1):
    c = ws.cell(row=i, column=4, value='=' + nm)
    c.font = Font(name=FONT, size=9, bold=(i == 1), color='404040')
    c.alignment = Alignment(horizontal='right')
    ws.merge_cells(start_row=i, start_column=4, end_row=i, end_column=6)
MESES = '"enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"'
ws['B6'] = '=par_Ciudad&", a "&DAY(inp_Fecha)&" de "&CHOOSE(MONTH(inp_Fecha),%s)&" de "&YEAR(inp_Fecha)' % MESES
ws['B6'].font = f_base
ws['F6'] = '="Folio: "&inp_Folio'
ws['F6'].font = f_bold
ws['F6'].alignment = Alignment(horizontal='right')
ws['B8'] = '=UPPER(inp_Cliente)'
ws['B8'].font = Font(name=FONT, size=11, bold=True)
ws['B9'] = '=IF(inp_Contacto="","","At\'n: "&inp_Contacto)'
ws['B9'].font = f_base
ws['B11'] = '="Por medio de la presente nos permitimos poner a su consideración la cotización de "&inp_Tipo&", con base en la siguiente información:"'
ws['B11'].font = f_base
ws['B11'].alignment = wrap
ws.merge_cells('B11:F11')
ws.row_dimensions[11].height = 28
header_bar(ws, 13, 2, 6, 'TÉRMINOS Y CONDICIONES')
info = [
    ('Tipo de operación:', '=inp_Tipo&" en "&inp_Moneda'),
    ('Arrendatario:', '=inp_Cliente'),
    ('Arrendador:', '=par_Arrendador'),
    ('Obligado solidario:', '=inp_Obligado'),
    ('Proveedor:', '=IF(inp_Proveedor="","Por definir",inp_Proveedor)'),
    ('Equipo:', '=inp_Equipo'),
    ('Valor del equipo:', '=TEXT(inp_Valor,"$#,##0.00")&" más IVA"'),
    ('Tipo de cambio (informativo):', '=IF(inp_Moneda="Moneda Nacional","No aplica",TEXT(inp_TC,"#,##0.0000"))'),
    ('Forma de pago de las rentas:', '="Mensual, "&LOWER(inp_Modalidad)&IF(inp_Modalidad="Anticipado"," (al inicio de cada mes)"," (al final de cada mes)")'),
]
r = 14
for label, f in info:
    ws.cell(row=r, column=2, value=label).font = f_bold
    c = ws.cell(row=r, column=3, value=f)
    c.font = f_base
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    r += 1
r += 1
TBL = r
ws.cell(row=r, column=2, value='CONCEPTO').font = f_hdr
for col in range(2, 7):
    ws.cell(row=r, column=col).fill = fill_hdr
    ws.cell(row=r, column=col).alignment = center
for k in range(1, 5):
    c = ws.cell(row=r, column=2 + k, value='=IF(INDEX(esc_Incluir,1,%d)="Sí","OPCIÓN "&%d,"")' % (k, k))
    c.font = f_hdr
ws.cell(row=r, column=2).alignment = Alignment(horizontal='left', vertical='center')
ws.row_dimensions[r].height = 20
r += 1


def carta_row(label, fn, fmt=MON, bold=False, sub=False, top=False):
    global r
    c0 = ws.cell(row=r, column=2, value=label)
    c0.font = f_bold if (bold or sub) else f_base
    if sub:
        for col in range(2, 7):
            ws.cell(row=r, column=col).fill = fill_sub
    for k in range(1, 5):
        c = ws.cell(row=r, column=2 + k)
        if fn:
            c.value = '=IF(INDEX(esc_Incluir,1,%d)="Sí",%s,"")' % (k, fn(k))
        c.number_format = fmt or 'General'
        c.font = f_bold if bold else f_base
        c.alignment = Alignment(horizontal='center')
        if top:
            c.border = Border(top=Side(style='thin', color='000000'))
    if top:
        c0.border = Border(top=Side(style='thin', color='000000'))
    r += 1


carta_row('Plazo básico', lambda k: 'INDEX(esc_Plazo,1,%d)&" meses"' % k, None, bold=True)
carta_row('Plazo sucesivo', lambda k: 'IF(INDEX(esc_Sucesivo,1,%d)>0,INDEX(esc_Sucesivo,1,%d)&IF(INDEX(esc_Sucesivo,1,%d)=1," mes"," meses"),"Opción de compra")' % (k, k, k), None)
carta_row('Monto a financiar (más IVA)', lambda k: cr('M', k))
carta_row('PAGO INICIAL (IVA incluido)', None, sub=True)
carta_row('Enganche', lambda k: '%s*(1+par_IVA)' % cr('eng', k))
carta_row('Comisión por apertura', lambda k: '%s*(1+par_IVA)' % cr('com', k))
carta_row('Depósito en garantía', lambda k: cr('dep', k))
carta_row('Renta proporcional', lambda k: '%s*(1+par_IVA)' % cr('rp', k))
carta_row('Seguro y GPS de contado', lambda k: 'ROUND((%s-%s-%s-%s)*(1+par_IVA),2)' % (cr('PIb', k), cr('eng', k), cr('com', k), cr('rp', k)))
carta_row('Total pago inicial', lambda k: cr('PI', k), bold=True, top=True)
carta_row('RENTAS MENSUALES', None, sub=True)
carta_row('Renta mensual (más IVA)', lambda k: cr('R', k))
carta_row('IVA', lambda k: cr('IVA1', k))
carta_row('Renta mensual con IVA', lambda k: '%s+%s' % (cr('R', k), cr('IVA1', k)), bold=True, top=True)
carta_row('Renta plazo sucesivo / opción de compra (más IVA)', lambda k: '%s+%s' % (cr('Rs', k), cr('OC', k)))
carta_row('Seguro y GPS financiados (incluidos en la renta)', lambda k: 'IF(%s+%s=0,"No aplica",TEXT(%s+%s,"$#,##0.00")&" + IVA")' % (cr('segf', k), cr('gpsf', k), cr('segf', k), cr('gpsf', k)), None)
for rr_ in range(TBL + 1, r):
    for col in range(2, 7):
        cc = ws.cell(row=rr_, column=col)
        if not cc.border.top.style:
            cc.border = Border(bottom=Side(style='hair', color='BFBFBF'))
r += 1
ws.cell(row=r, column=2, value='Cifras en la moneda de la operación. Las rentas causan IVA a la tasa vigente. Sujeto a aprobación de crédito.').font = f_note
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
r += 2


def texto(label_row_cfg, est_chars):
    global r
    cfgrow = TXT_ROW0 + label_row_cfg
    lab = ws.cell(row=r, column=2, value="=Configuracion!$B$%d" % cfgrow)
    lab.font = f_bold
    lab.alignment = wrap
    c = ws.cell(row=r, column=3, value='=SUBSTITUTE(SUBSTITUTE(Configuracion!$C$%d,"{ARRENDADOR}",par_Arrendador),"{VIGENCIA}",par_Vigencia)' % cfgrow)
    c.font = Font(name=FONT, size=9)
    c.alignment = Alignment(wrap_text=True, vertical='top', horizontal='justify')
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    lines = est_chars // 95 + 1
    ws.row_dimensions[r].height = max(14, 12 * lines + 4)
    r += 1


for i, (label, txt) in enumerate(TEXTOS[:-1]):
    if i == 2:
        ws.cell(row=r, column=2, value='Esta propuesta se basa en los siguientes términos y condiciones adicionales:').font = f_bold
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        r += 1
    texto(i, len(txt) + 20)
r += 1
c = ws.cell(row=r, column=2, value="=Configuracion!$C$%d" % (TXT_ROW0 + len(TEXTOS) - 1))
c.font = f_base
c.alignment = wrap
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
ws.row_dimensions[r].height = 28
r += 3
ws.cell(row=r, column=2, value='A T E N T A M E N T E').font = f_bold
ws.cell(row=r, column=4, value='PROPUESTA ACEPTADA POR:').font = f_bold
r += 3
SIG = r
ws.cell(row=r, column=2).border = Border(bottom=Side(style='thin', color='000000'))
for lab in ['Nombre:', 'Puesto:', 'Fecha:', 'Firma:']:
    ws.cell(row=r, column=4, value=lab).font = f_base
    for col in (5, 6):
        ws.cell(row=r, column=col).border = Border(bottom=Side(style='thin', color='000000'))
    r += 1
ws.cell(row=SIG + 1, column=2, value='=inp_Promotor').font = f_bold
ws.cell(row=SIG + 2, column=2, value='="Promoción – "&par_Comercial').font = f_base
ws.cell(row=SIG + 3, column=2, value='=IF(par_Contacto="","",par_Contacto)').font = f_note
LAST_CARTA = r
ws.print_area = 'A1:G%d' % LAST_CARTA
ws.page_setup.orientation = 'portrait'
ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_margins = page.PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.6)
ws.oddFooter.center.text = 'Página &P de &N'
ws.oddFooter.center.size = 8

# ==================================================================== TABLA DE PAGOS
ws = ws_tab
for col, w in zip('ABCDEFGHI', [2, 7, 13, 22, 17, 15, 17, 18, 2]):
    ws.column_dimensions[col].width = w
logo(ws, 'B1', 55)
ws['D1'] = 'TABLA DE PAGOS'
ws['D1'].font = f_title
ws.merge_cells('D1:H1')
ws['D2'] = '=par_Arrendador'
ws['D2'].font = f_sub
ws.merge_cells('D2:H2')
lab = [('Cliente:', '=inp_Cliente'), ('Equipo:', '=inp_Equipo'), ('Folio:', '=inp_Folio')]
for i, (a, b) in enumerate(lab):
    ws.cell(row=4 + i, column=2, value=a).font = f_bold
    ws.cell(row=4 + i, column=4, value=b).font = f_base
    ws.merge_cells(start_row=4 + i, start_column=4, end_row=4 + i, end_column=8)
ws['B7'] = 'Escenario a mostrar:'
ws['B7'].font = f_bold
ws['F7'] = 'Elija 1, 2, 3 ó 4'
ws['F7'].font = f_note
ws.merge_cells('B7:C7')
inp(ws['D7'], 1, '0')
name('inp_EscTabla', absref(ws.title, 'D7'))
dv_list(ws, '"1,2,3,4"', 'D7')
CH = lambda ref: 'CHOOSE(inp_EscTabla,%s)' % ','.join("'Corrida %d'!%s" % (k, ref) for k in range(1, 5))
ws['B8'] = '="Plazo: "&%s&" meses + "&%s&" sucesivo  |  Tasa anual: "&TEXT(%s,"0.00%%")&"  |  Pago "&LOWER(inp_Modalidad)&"  |  "&inp_Moneda' % (CH('$C$5'), CH('$C$6'), CH('$C$8'))
ws['B8'].font = f_base
ws.merge_cells('B8:H8')
hdr = ['No.', 'Fecha', 'Concepto', 'Importe sin IVA', 'IVA', 'Total con IVA', 'Acumulado con IVA']
for j, h in enumerate(hdr):
    c = ws.cell(row=10, column=2 + j, value=h)
    c.font = f_hdr
    c.fill = fill_hdr
    c.alignment = center
# pago inicial
r = 11
ws.cell(row=r, column=2, value=0)
ws.cell(row=r, column=3, value='=inp_FechaFirma').number_format = FECHA
ws.cell(row=r, column=4, value='Pago inicial')
ws.cell(row=r, column=5, value='=%s+%s' % (CH('$I$17'), CH('$I$19'))).number_format = MON
ws.cell(row=r, column=6, value='=%s' % CH('$I$18')).number_format = MON
ws.cell(row=r, column=7, value='=E11+F11').number_format = MON
ws.cell(row=r, column=8, value='=G11').number_format = MON
for col in range(2, 9):
    ws.cell(row=r, column=col).fill = fill_gris
    ws.cell(row=r, column=col).font = f_base
for kk in range(1, NROWS + 1):
    r = 11 + kk
    src = FIRST + kk - 1
    ws.cell(row=r, column=2, value='=IF(%s="","",%d)' % (CH('$B$%d' % src), kk))
    ws.cell(row=r, column=3, value='=IF(B%d="","",%s)' % (r, CH('$D$%d' % src))).number_format = FECHA
    ws.cell(row=r, column=4, value='=IF(B%d="","",IF(%s="Básico","Renta mensual",IF(%s="Sucesivo","Renta plazo sucesivo","Opción de compra")))' % (r, CH('$B$%d' % src), CH('$B$%d' % src)))
    ws.cell(row=r, column=5, value='=IF(B%d="","",%s)' % (r, CH('$F$%d' % src))).number_format = MON
    ws.cell(row=r, column=6, value='=IF(B%d="","",%s)' % (r, CH('$J$%d' % src))).number_format = MON
    ws.cell(row=r, column=7, value='=IF(B%d="","",E%d+F%d)' % (r, r, r)).number_format = MON
    ws.cell(row=r, column=8, value='=IF(B%d="","",H%d+G%d)' % (r, r - 1, r)).number_format = MON
    for col in range(2, 9):
        ws.cell(row=r, column=col).font = f_base
        ws.cell(row=r, column=col).alignment = Alignment(horizontal='center') if col <= 4 else Alignment()
TOTR = 11 + NROWS + 1
ws.cell(row=TOTR, column=4, value='TOTAL').font = f_bold
for col, L_ in ((5, 'E'), (6, 'F'), (7, 'G')):
    c = ws.cell(row=TOTR, column=col, value='=SUM(%s11:%s%d)' % (L_, L_, TOTR - 1))
    c.number_format = MON
    c.font = f_bold
    c.border = Border(top=Side(style='thin', color='000000'))
ws.cell(row=TOTR + 1, column=2, value='El depósito en garantía se incluye en el pago inicial y se aplica en el último pago. Cifras sujetas a aprobación de crédito.').font = f_note
ws.merge_cells(start_row=TOTR + 1, start_column=2, end_row=TOTR + 1, end_column=8)
ws.conditional_formatting.add('B12:H%d' % (TOTR - 1), FormulaRule(formula=['AND($B12<>"",MOD($B12,2)=0)'], fill=PatternFill('solid', fgColor='F7F4FB')))
ws.freeze_panes = 'B11'
ws.print_area = 'A1:I%d' % (TOTR + 1)
ws.print_title_rows = '10:10'
ws.page_setup.orientation = 'portrait'
ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.oddFooter.center.text = 'Página &P de &N'

# ==================================================================== SENSIBILIDAD
ws = ws_sen
for col, w in zip('ABCDEFGHIJ', [2, 22, 15, 15, 15, 15, 15, 15, 15, 2]):
    ws.column_dimensions[col].width = w
ws['B1'] = 'ANÁLISIS DE SENSIBILIDAD – RENTA MENSUAL'
ws['B1'].font = f_title
ws['B2'] = 'Usa el valor del equipo, enganche, residual, comisión y financiados del escenario elegido. Cambie las tasas amarillas libremente.'
ws['B2'].font = f_sub
ws['B4'] = 'Escenario base:'
ws['B4'].font = f_bold
inp(ws['C4'], 1, '0')
dv_list(ws, '"1,2,3,4"', 'C4')
CHs = lambda ref: 'CHOOSE($C$4,%s)' % ','.join("'Corrida %d'!%s" % (k, ref) for k in range(1, 5))
ws['B5'] = 'Monto a financiar:'
ws['C5'] = '=%s' % CHs('$C$19')
ws['C5'].number_format = MON
ws['B6'] = 'Valor residual:'
ws['C6'] = '=%s' % CHs('$C$21')
ws['C6'].number_format = MON
ws['B7'] = 'Tipo de pago:'
ws['C7'] = '=%s' % CHs('$C$7')
ws['D7'] = '=IF(C7=1,"anticipado","vencido")'
for c in ('B5', 'B6', 'B7'):
    ws[c].font = f_bold
for c in ('C5', 'C6', 'C7', 'D7'):
    ws[c].font = Font(name=FONT, size=10, color='008000')
header_bar(ws, 9, 2, 9, 'RENTA MENSUAL SIN IVA  (filas: plazo en meses · columnas: tasa anual)')
ws['B10'] = 'Plazo \\ Tasa'
ws['B10'].font = f_bold
steps = [-0.03, -0.02, -0.01, 0, 0.01, 0.02, 0.03]
for j, s in enumerate(steps):
    inp(ws.cell(row=10, column=3 + j), '=ROUND(%s+(%s),4)' % (CHs('$C$8'), s), PCT)
plazos = [12, 18, 24, 30, 36, 42, 48, 54, 60, 72, 84]
for i, n in enumerate(plazos):
    rr = 11 + i
    inp(ws.cell(row=rr, column=2), n, '0" meses"')
    for j in range(7):
        col = L(3 + j)
        c = ws.cell(row=rr, column=3 + j, value='=IFERROR(PMT(%s$10/12,$B%d,-$C$5,$C$6,$C$7),"")' % (col, rr))
        c.number_format = MON
        c.font = f_base
        c.border = box
R2 = 11 + len(plazos) + 1
header_bar(ws, R2, 2, 9, 'RENTA MENSUAL CON IVA')
for i, n in enumerate(plazos):
    rr = R2 + 1 + i
    c = ws.cell(row=rr, column=2, value='=B%d' % (11 + i))
    c.number_format = '0" meses"'
    c.font = f_bold
    for j in range(7):
        col = L(3 + j)
        c = ws.cell(row=rr, column=3 + j, value='=IFERROR(%s%d*(1+par_IVA),"")' % (col, 11 + i))
        c.number_format = MON
        c.font = f_base
        c.border = box
ws.conditional_formatting.add('C11:I%d' % (10 + len(plazos)), FormulaRule(formula=['AND($B11=%s,C$10=%s)' % (CHs('$C$5'), CHs('$C$8'))], fill=PatternFill('solid', fgColor='C6EFCE')))

# ==================================================================== HISTORIAL
ws = ws_his
hist_cols = ['Folio', 'Fecha', 'Cliente', 'Equipo', 'Promotor', 'Moneda', 'Valor equipo', 'Plazos',
             'Renta c/IVA Esc.1', 'Renta c/IVA Esc.2', 'Renta c/IVA Esc.3', 'Renta c/IVA Esc.4',
             'CM % Esc.1', 'CM % Esc.2', 'CM % Esc.3', 'CM % Esc.4', 'Estatus', 'Comentarios', 'Guardado por', 'Guardado el']
widths = [12, 11, 30, 26, 22, 14, 15, 14, 15, 15, 15, 15, 10, 10, 10, 10, 15, 30, 16, 17]
ws['A1'] = 'HISTORIAL DE COTIZACIONES'
ws['A1'].font = f_title
ws['A2'] = 'La macro "Guardar en historial" agrega una fila por folio. Para reabrir una cotización seleccione su fila y use "Cargar del historial". Las columnas a la derecha guardan todos los datos capturados.'
ws['A2'].font = f_sub
for j, (h, w) in enumerate(zip(hist_cols, widths), start=1):
    c = ws.cell(row=4, column=j, value=h)
    c.font = f_hdr
    c.fill = fill_hdr
    c.alignment = center
    ws.column_dimensions[L(j)].width = w
ws.row_dimensions[4].height = 30
ws.freeze_panes = 'B5'
ws.auto_filter.ref = 'A4:T4'
dv_list(ws, '=lst_Estatus', 'Q5:Q2000')
name('hist_Header', absref(ws.title, 'A4'))

# ==================================================================== INICIO
ws = ws_ini
for col, w in zip('ABCDEF', [2, 34, 70, 3, 3, 3]):
    ws.column_dimensions[col].width = w
logo(ws, 'B1', 70)
ws['C2'] = 'COTIZADOR DE ARRENDAMIENTO'
ws['C2'].font = Font(name=FONT, size=22, bold=True, color=MORADO)
ws['C3'] = 'Arrendamiento puro y financiero · 4 escenarios · cash margin · TIR · CAT · carta y tabla de pagos'
ws['C3'].font = f_sub
ws.row_dimensions[2].height = 30
r = 6
header_bar(ws, r, 2, 3, 'PASOS PARA COTIZAR')
pasos = [
    ('1. Nueva cotización', 'Cinta "Cotizador Arrendamiento" → Nueva cotización (o Ctrl+Shift+N). Limpia la captura y asigna folio.'),
    ('2. Capturar datos', 'En la hoja Cotizador llene solo las celdas AMARILLAS (texto azul): cliente, equipo, valor, fechas y los 4 escenarios.'),
    ('3. Revisar rentabilidad', 'Sección 5 del Cotizador: cash margin, TIR y dictamen (verde = cumple). Si no cumple use "Aplicar tasa mínima".'),
    ('4. Revisar la carta', 'Hoja Carta Cotizacion: lista para imprimir. Los escenarios con "Incluir = No" no aparecen.'),
    ('5. Entregar', 'Carta en PDF (Ctrl+Shift+P), Tabla de pagos en PDF o Enviar por correo (Outlook).'),
    ('6. Guardar', 'Guardar en historial (Ctrl+Shift+G). Desde Historial puede volver a cargar cualquier folio.'),
]
r += 1
for a, b in pasos:
    ws.cell(row=r, column=2, value=a).font = f_bold
    ws.cell(row=r, column=3, value=b).font = f_base
    ws.cell(row=r, column=3).alignment = wrap
    ws.row_dimensions[r].height = 30
    r += 1
r += 1
header_bar(ws, r, 2, 3, 'IR A…')
r += 1
for sh, desc in [('Cotizador', 'Captura y resultados'), ('Carta Cotizacion', 'Carta para el cliente'),
                 ('Tabla de Pagos', 'Calendario de pagos con IVA'), ('Sensibilidad', 'Matriz de rentas por plazo y tasa'),
                 ('Historial', 'Registro de cotizaciones'), ('Corrida 1', 'Corrida financiera detallada (también 2, 3 y 4)'),
                 ('Configuracion', 'Empresa, IVA, fondeo, promotores, textos (administrador)')]:
    c = ws.cell(row=r, column=2, value=sh)
    c.hyperlink = "#'%s'!A1" % sh
    c.font = Font(name=FONT, size=10, color='0563C1', underline='single')
    ws.cell(row=r, column=3, value=desc).font = f_base
    r += 1
r += 1
header_bar(ws, r, 2, 3, 'CÓDIGO DE COLORES')
r += 1
leyenda = [('Celda amarilla, texto azul', 'Dato que usted captura', fill_in, f_input),
           ('Texto verde', 'Resultado tomado de otra hoja (no editar)', None, f_link),
           ('Texto negro', 'Fórmula (no editar)', None, f_base)]
for a, b, fl, fo in leyenda:
    c = ws.cell(row=r, column=2, value=a)
    c.font = fo
    if fl:
        c.fill = fl
    ws.cell(row=r, column=3, value=b).font = f_base
    r += 1
r += 1
header_bar(ws, r, 2, 3, 'MACROS DISPONIBLES (cinta "Cotizador Arrendamiento" o Alt+F8)')
r += 1
macros = [
    ('NuevaCotizacion  (Ctrl+Shift+N)', 'Limpia la captura, pone valores por defecto y el siguiente folio.'),
    ('GuardarCotizacion  (Ctrl+Shift+G)', 'Agrega/actualiza el folio en Historial con todos los datos capturados.'),
    ('CargarCotizacion', 'Recupera una cotización del Historial (fila seleccionada o folio).'),
    ('ExportarCartaPDF  (Ctrl+Shift+P)', 'Guarda la carta en PDF: Cotizacion_<folio>_<cliente>.pdf'),
    ('ExportarTablaPDF', 'Guarda la tabla de pagos del escenario elegido en PDF.'),
    ('EnviarPorCorreo', 'Genera el PDF y abre un correo de Outlook al cliente con el PDF adjunto.'),
    ('AplicarTasaMinima', 'Pone en el escenario la tasa mínima que cumple el cash margin objetivo.'),
    ('TasaParaRentaDeseada', 'Calcula y aplica la tasa que da la renta que pide el cliente.'),
    ('CopiarEscenario1', 'Copia tasas, enganche, residual, comisión, seguros, etc. del escenario 1 a los demás.'),
    ('ProtegerHojas / DesprotegerHojas', 'Protege (sin contraseña) las fórmulas para evitar borrados accidentales.'),
]
for a, b in macros:
    ws.cell(row=r, column=2, value=a).font = f_bold
    ws.cell(row=r, column=3, value=b).font = f_base
    ws.cell(row=r, column=3).alignment = wrap
    r += 1
r += 1
header_bar(ws, r, 2, 3, 'MODELO MATEMÁTICO')
r += 1
modelo = [
    ('Renta básica', 'R = PMT(i, n, −M, VR, tipo), con i = tasa anual/12, M = valor − enganche + seguro/GPS financiados, VR = residual × valor, tipo = 1 anticipado / 0 vencido.'),
    ('Renta sucesiva', 'Rs = PMT(i2, m, −PF, 0, tipo). Si m = 0 el pago final se cobra como opción de compra única.'),
    ('Amortización', 'Interés_t = (Saldo_t − tipo × Pago_t) × i ; Capital = Pago − Interés. El saldo al final del plazo básico es exactamente VR y termina en 0.'),
    ('Cash margin', 'CM = flujo en firma + Σ flujo_t × (1+fondeo/360)^(−días) × (1+fondeo/12)^(−t). Flujo en firma = −(valor + financiados) + enganche + comisión + renta proporcional + depósito − comisión banco. El depósito se devuelve en el último pago.'),
    ('Tasa mínima', 'CM es lineal en R: CM(R) = A + B·R. Renta mínima = (CM objetivo × M − A)/B ; tasa mínima = RATE(n, Rmín, −M, VR, tipo) × 12. Sin tanteos ni Buscar objetivo.'),
    ('TIR y CAT', 'TIR = XIRR de los flujos del arrendador con fechas reales. CAT informativo = XIRR de los flujos del cliente sin IVA (no oficial).'),
]
for a, b in modelo:
    ws.cell(row=r, column=2, value=a).font = f_bold
    ws.cell(row=r, column=3, value=b).font = f_base
    ws.cell(row=r, column=3).alignment = wrap
    ws.row_dimensions[r].height = 42
    r += 1
r += 1
ws.cell(row=r, column=2, value='Versión 2.0 – octubre 2026').font = f_note

# ==================================================================== IMPRESIÓN
for w_, orient in ((ws_his, 'landscape'), (ws_sen, 'landscape'), (ws_ini, 'portrait'), (ws_cfg, 'landscape')):
    w_.page_setup.orientation = orient
    w_.page_setup.paperSize = w_.PAPERSIZE_LETTER
    w_.page_setup.fitToWidth = 1
    w_.page_setup.fitToHeight = 0
    w_.sheet_properties.pageSetUpPr.fitToPage = True
ws_his.print_title_rows = '4:4'
ws_sen.print_area = 'A1:J%d' % (R2 + len(plazos))

# ==================================================================== PROTECCIÓN
for w in [ws_cot, ws_car, ws_tab, ws_sen] + ws_cr:
    w.protection.sheet = True
    w.protection.formatColumns = False
    w.protection.formatRows = False
    w.protection.selectLockedCells = False
    w.protection.selectUnlockedCells = False

wb.active = 1
wb.save(OUT)
import json
json.dump({'P': P, 'names': names, 'FIRST': FIRST, 'LAST': LAST}, open('layout.json', 'w'), indent=1, ensure_ascii=False)
print('ok', OUT)
