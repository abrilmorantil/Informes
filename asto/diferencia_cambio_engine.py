"""
Motor de Ajuste por Conversión / Diferencia de Cambio (bimonetario ARS/USD).
Cliente de referencia: Southern Copper Argentina S.R.L. (SCA).

Reproduce el asiento "Ajuste por Conversión" que hoy se arma a mano en
Asto_dif_cambio_<periodo>.xlsx, a partir del Balance de Sumas y Saldos (SyS)
bimonetario y de dos tipos de cambio de cierre.

Validado contra el asiento real Nº 2630 (06-2026): 89/89 líneas materiales
coinciden al centavo; el residuo sub-centavo cae en la cuenta de balanceo.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
import re

# ---------------------------------------------------------------------------
# CONFIGURACIÓN PERSISTENTE (se mantiene entre períodos, cambia rara vez)
# ---------------------------------------------------------------------------

# Cuentas NO monetarias: se re-expresan a costo histórico, NO se traducen.
# Su diferencia peso/USD es real pero NO debe ajustarse -> exclusión por criterio
# contable (no se auto-excluye por el filtro de materialidad).
NON_MONETARY_PREFIXES = ("1240", "1250")          # Bienes de Uso + Dep. Acum., Cargos Diferidos
NON_MONETARY_EXACT    = {"114010016",             # Impuesto Crédito Diferido
                         "114050005",             # Seguros a Devengar (pago anticipado)
                         "211050000"}             # Previsión AID (impuesto diferido pasivo)

# Subdiario de proveedores: se calcula SIN redondear el cociente intermedio
# (replica hoja "Proveedores": C = B/TC, E = C - D). El resto de las cuentas
# usa REDONDEAR(peso/TC, 2) (replica hoja "TRADUCC": E = ROUND(C/TC,2), I = E - G).
SUPPLIER_PREFIX = "21101"

# Cuenta de balanceo (resultado): absorbe el neto del ajuste y el residuo de redondeo.
CUENTA_DIFERENCIA_CAMBIO = "423050000"
DENOM_DIFERENCIA_CAMBIO  = "Diferencia de Cambio USD"

# Umbral de materialidad para publicar una línea (en USD).
MATERIALIDAD = 0.005

# Cuentas monetarias YA confirmadas por el usuario (allowlist): no se re-marcan como "nuevas".
CONFIRMED_MONETARY: set[str] = set()

# Salvaguarda: se marca sospechosa la línea que cumple AMBAS condiciones
# (desproporción + materialidad). Calibrado con datos reales: las cuentas no
# monetarias mal incluidas dan ratio ~0.95, las monetarias legítimas ~0.05.
UMBRAL_RATIO = 0.30    # |ajuste| > 30% del saldo USD en libros  (separa 0.05 de 0.95)
UMBRAL_ABS   = 500.0   # Y además |ajuste| > 500 USD (evita falsos positivos de saldo chico)

CONCEPTO = "Ajuste por Conversión"


# ---------------------------------------------------------------------------
# MODELO DE DATOS
# ---------------------------------------------------------------------------
@dataclass
class CuentaSaldo:
    codigo: str          # 9 dígitos, ej "111010001"
    denominacion: str
    seccion: str         # ACTIVO | PASIVO | PATRIMONIO NETO | RESULTADOS
    saldo_pesos: float   # columna "Saldo ($)"
    saldo_usd: float     # columna "Saldo (u$s)" -> saldo actualmente en libros USD

@dataclass
class LineaAjuste:
    codigo: str
    denominacion: str
    seccion: str
    tc_aplicado: float
    usd_teorico: float   # peso / TC (redondeado o no según tipo de cuenta)
    usd_libros: float
    ajuste_usd: float    # usd_teorico - usd_libros  (con signo)

@dataclass
class Parametros:
    periodo_fin: date         # fecha de cierre, ej date(2026, 6, 30)
    tc_compra: float          # T.C. COMPRA de cierre -> cuentas del ACTIVO
    tc_venta: float           # T.C. VENTA  de cierre -> cuentas del PASIVO
    numero_asiento: int = 1   # opcional, para el importador
    concepto: str = CONCEPTO


# ---------------------------------------------------------------------------
# 1) PARSEO DEL SyS (.xls Crystal Reports, bimonetario)
# ---------------------------------------------------------------------------
def parse_sys_xls(path: str) -> list[CuentaSaldo]:
    """Lee el Balance de Sumas y Saldos bimonetario.
    Columnas relevantes (0-based): A=0 (Cuenta - Denominación),
    P=15 (Saldo $), AB=27 (Saldo u$s). Secciones marcadas por texto en col A."""
    import xlrd
    wb = xlrd.open_workbook(path)
    s = wb.sheet_by_index(0)
    COL_CUENTA, COL_SALDO_PESOS, COL_SALDO_USD = 0, 15, 27
    SECCIONES = {"ACTIVO", "PASIVO", "PATRIMONIO NETO", "RESULTADOS"}
    seccion = None
    out: list[CuentaSaldo] = []
    for r in range(s.nrows):
        a = s.cell_value(r, COL_CUENTA)
        if not isinstance(a, str):
            continue
        a = a.strip()
        if a in SECCIONES:
            seccion = a
            continue
        m = re.match(r"^(\d{9})\s*-\s*(.*)$", a)
        if m and seccion:
            out.append(CuentaSaldo(
                codigo=m.group(1),
                denominacion=m.group(2).strip(),
                seccion=seccion,
                saldo_pesos=_num(s.cell_value(r, COL_SALDO_PESOS)),
                saldo_usd=_num(s.cell_value(r, COL_SALDO_USD)),
            ))
    return out

def _num(v) -> float:
    try:
        return float(v) if v not in ("", None) else 0.0
    except (TypeError, ValueError):
        return 0.0


# ---------------------------------------------------------------------------
# 2) CLASIFICACIÓN Y CÁLCULO
# ---------------------------------------------------------------------------
def es_monetaria(c: CuentaSaldo) -> bool:
    if c.seccion not in ("ACTIVO", "PASIVO"):
        return False
    if c.codigo[:4] in NON_MONETARY_PREFIXES:
        return False
    if c.codigo in NON_MONETARY_EXACT:
        return False
    return True

def tc_de(c: CuentaSaldo, p: Parametros) -> float:
    # ACTIVO -> T.C. COMPRA ; PASIVO -> T.C. VENTA
    return p.tc_compra if c.seccion == "ACTIVO" else p.tc_venta

def calcular_lineas(cuentas: list[CuentaSaldo], p: Parametros) -> list[LineaAjuste]:
    lineas: list[LineaAjuste] = []
    for c in cuentas:
        if not es_monetaria(c):
            continue
        tc = tc_de(c, p)
        if c.codigo.startswith(SUPPLIER_PREFIX):
            usd_teorico = c.saldo_pesos / tc            # proveedores: sin redondear
        else:
            usd_teorico = round(c.saldo_pesos / tc, 2)  # generales: redondeado
        ajuste = usd_teorico - c.saldo_usd
        if abs(round(ajuste, 2)) >= MATERIALIDAD:
            lineas.append(LineaAjuste(c.codigo, c.denominacion, c.seccion,
                                      tc, usd_teorico, c.saldo_usd, ajuste))
    return lineas

def motivo_sospecha(l: LineaAjuste) -> str | None:
    """Devuelve el motivo si la línea debe pasar por revisión del usuario, o None."""
    if l.codigo in CONFIRMED_MONETARY:
        return None
    desproporcion = (abs(l.usd_libros) > 1e-9
                     and abs(l.ajuste_usd) > UMBRAL_RATIO * abs(l.usd_libros))
    material = abs(l.ajuste_usd) > UMBRAL_ABS
    # Ambas: una no monetaria mal incluida es a la vez desproporcionada y grande.
    if desproporcion and material:
        return "desproporcion"
    return None


def calcular_con_revision(cuentas: list[CuentaSaldo], p: Parametros):
    """Fase 1: NO genera el importador. Separa líneas OK de las que requieren
    confirmación del usuario (posibles no monetarias mal incluidas)."""
    lineas = calcular_lineas(cuentas, p)
    ok, revisar = [], []
    for l in lineas:
        m = motivo_sospecha(l)
        (revisar if m else ok).append((l, m) if m else l)
    return {"lineas_ok": ok, "lineas_a_revisar": revisar}  # revisar = [(linea, motivo), ...]


def aplicar_decisiones(lineas_a_revisar, decisiones: dict[str, str]):
    """Fase 2 (tras la UI): decisiones = {codigo: 'incluir' | 'excluir'}.
    Devuelve las líneas incluidas y la config a persistir para no repreguntar."""
    incluidas, a_excluir, a_confirmar = [], [], []
    for l, _motivo in lineas_a_revisar:
        d = decisiones.get(l.codigo, "excluir")  # por defecto, no incluir sin confirmación
        if d == "incluir":
            incluidas.append(l); a_confirmar.append(l.codigo)
        else:
            a_excluir.append(l.codigo)
    nueva_config = {"non_monetary_exact_add": a_excluir,
                    "confirmed_monetary_add": a_confirmar}
    return incluidas, nueva_config


def armar_asiento(lineas: list[LineaAjuste]) -> list[LineaAjuste]:
    """Agrega la línea de balanceo (Diferencia de Cambio) redondeando cada
    ajuste a 2 decimales; la cuenta de resultado absorbe el residuo para que
    el asiento cierre exactamente en 0."""
    publicadas = [
        LineaAjuste(l.codigo, l.denominacion, l.seccion, l.tc_aplicado,
                    l.usd_teorico, l.usd_libros, round(l.ajuste_usd, 2))
        for l in lineas
    ]
    neto = round(sum(l.ajuste_usd for l in publicadas), 2)
    publicadas.append(LineaAjuste(
        CUENTA_DIFERENCIA_CAMBIO, DENOM_DIFERENCIA_CAMBIO, "RESULTADOS",
        0.0, 0.0, 0.0, round(-neto, 2)))
    return publicadas


# ---------------------------------------------------------------------------
# 3) SALIDA: archivo importador Onvio (formato "Asientos", igual a 0.xls)
# ---------------------------------------------------------------------------
IMPORT_HEADERS = [
    "Número de asiento", "Número de Pase", "Fecha", "Concepto",
    "Código de cuenta", "Importe en moneda local",
    "Importe en moneda ext.present.", "Leyenda",
    "Código de centro de costos", "Porcentaje de distribución",
    "Imp.mon.local dist.C.Costos", "Imp.mon.present.dist.C.Costos",
]

def _excel_serial(d: date) -> int:
    return (datetime(d.year, d.month, d.day) - datetime(1899, 12, 30)).days

def escribir_importador_xls(lineas: list[LineaAjuste], p: Parametros, path: str):
    """Genera el .xls con las columnas exactas del importador.
       F (moneda local / pesos) = 0 ; G (moneda ext. present. / USD) = ajuste con signo
       (positivo = Debe MEP, negativo = Haber MEP)."""
    import xlwt
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Asientos")
    for c, h in enumerate(IMPORT_HEADERS):
        ws.write(0, c, h)
    date_fmt = xlwt.easyxf(num_format_str="M/D/YYYY")
    serial = _excel_serial(p.periodo_fin)
    for i, l in enumerate(lineas, start=1):
        row = i
        ws.write(row, 0, p.numero_asiento)      # A Número de asiento
        ws.write(row, 1, i)                      # B Número de Pase
        ws.write(row, 2, serial, date_fmt)       # C Fecha
        ws.write(row, 3, p.concepto)             # D Concepto
        ws.write(row, 4, l.codigo)               # E Código de cuenta (9 dígitos)
        ws.write(row, 5, 0.0)                     # F Importe moneda local (pesos = 0)
        ws.write(row, 6, round(l.ajuste_usd, 2)) # G Importe moneda ext. present. (USD firmado)
        # H..L en blanco
    wb.save(path)


# ---------------------------------------------------------------------------
# Orquestador
# ---------------------------------------------------------------------------
def generar(sys_path: str, p: Parametros, importador_path: str):
    cuentas = parse_sys_xls(sys_path)
    lineas = calcular_lineas(cuentas, p)
    asiento = armar_asiento(lineas)
    escribir_importador_xls(asiento, p, importador_path)
    return asiento
