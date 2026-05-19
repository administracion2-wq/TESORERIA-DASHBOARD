"""
=============================================================
  ACTUALIZADOR DE DASHBOARD - CMA Paraguay S.A.
=============================================================
  Uso:
    1. Guardá el nuevo reporte SAP (.xlsx) en esta misma carpeta
    2. Hacé doble click en este archivo  (o ejecutá: python actualizar_dashboard.py)
    3. ¡Listo! Los 4 dashboards quedan actualizados automáticamente

  Requisitos: Python 3.8+ con openpyxl instalado
    pip install openpyxl
=============================================================
"""

import os, sys, re, json, glob
from datetime import datetime, date

# ─────────────────────────────────────────────────────────────
#  CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────
CARPETA       = os.path.dirname(os.path.abspath(__file__))
CARPETA_DATA  = os.path.join(CARPETA, "data")   # Excel files live here (GitHub mode)
DASH_OPS      = os.path.join(CARPETA, "Dashboard_Operaciones.html")
DASH_BANCOS   = os.path.join(CARPETA, "Dashboard_Bancos.html")
DASH_CLIENTES = os.path.join(CARPETA, "Dashboard_Clientes.html")
DASH_RESUMEN  = os.path.join(CARPETA, "Dashboard_Resumen.html")
DASH_FLUJO    = os.path.join(CARPETA, "Dashboard_Flujo.html")

AGING_BUCKETS = ["Por vencer", "1-30 días", "31-60 días",
                 "61-90 días", "91-180 días", "180+ días"]

HISTORIAL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "historial_deuda.json")
TC_REF = 7600   # Tipo de cambio de referencia para el historial (se puede ajustar)

# Clientes declarados en quiebra / incobrables — excluidos del flujo de caja
# y marcados visualmente en el dashboard de clientes
CLIENTES_INCOBRABLES = ["CAPITAL TRADE"]

# ─────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────
def log(msg):
    print(f"  {msg}")

def parse_monto(s):
    """Convierte 'GS 9.600.000' → (9600000.0, 'GS')
                  'USD 3.650,000' → (3650.0,   'USD')
       También acepta números crudos (float/int) → (valor, 'GS')"""
    if s is None or s == "":
        return (0.0, "GS")
    # Número crudo de Excel (openpyxl lo lee como float/int)
    if isinstance(s, (int, float)):
        return (abs(float(s)), "GS")
    s = str(s).strip()
    if not s or s.upper() in ("NONE", "0", "0.0"):
        return (0.0, "GS")
    mon = "USD" if "USD" in s.upper() else "GS"
    # Remover prefijos de moneda
    num = s.upper().replace("USD","").replace("GS","").strip()
    # Detectar formato: si tiene punto Y coma, el punto es miles y la coma es decimal
    if "." in num and "," in num:
        num = num.replace(".","").replace(",",".")
    elif "," in num and "." not in num:
        # Solo coma → separador decimal
        num = num.replace(",",".")
    elif "." in num and "," not in num:
        # Solo punto → puede ser decimal o miles; si >3 dígitos después → miles
        parts = num.split(".")
        if len(parts[-1]) == 3 and len(parts) > 1:
            num = num.replace(".","")  # miles
        # sino dejarlo como está (decimal)
    try:
        return (abs(float(num)), mon)
    except:
        return (0.0, mon)


def find_col(headers, candidates):
    """Busca índice de columna por nombre (parcial, sin distinción mayúsc.)"""
    headers_up = [h.upper().strip() for h in headers]
    for cand in candidates:
        for i, h in enumerate(headers_up):
            if cand.upper() in h:
                return i
    return None

def aging_ap(dias):
    d = int(dias or 0)
    if d <= 0:  return "Por vencer"
    if d <= 30: return "1-30 días"
    if d <= 60: return "31-60 días"
    if d <= 90: return "61-90 días"
    if d <= 180:return "91-180 días"
    return "180+ días"

def aging_ar(dias, estado):
    if estado == "COBRADO":   return "Cobrado"
    if estado == "A VENCER":  return "A Vencer"
    d = int(dias or 0)
    if d <= 30: return "1-30 días"
    if d <= 60: return "31-60 días"
    if d <= 90: return "61-90 días"
    if d <= 180:return "91-180 días"
    return "180+ días"

def rating_ar(dias, estado):
    if estado == "COBRADO":  return "COBRADO"
    if estado == "A VENCER": return "A VENCER"
    d = int(dias or 0)
    if d <= 30:  return "Riesgo Bajo"
    if d <= 60:  return "Riesgo Medio"
    if d <= 90:  return "Riesgo Alto"
    return "Crítico"

def fecha_ym(dt):
    """Retorna 'YYYY-MM' desde un datetime o string"""
    if isinstance(dt, datetime): return dt.strftime("%Y-%m")
    if isinstance(dt, date):     return dt.strftime("%Y-%m")
    if isinstance(dt, str) and len(dt) >= 7: return dt[:7]
    return "2026-01"

MESES = {1:"enero",2:"febrero",3:"marzo",4:"abril",5:"mayo",
         6:"junio",7:"julio",8:"agosto",9:"septiembre",
         10:"octubre",11:"noviembre",12:"diciembre"}

def replace_data_block(html, var_name, new_json_str):
    """Reemplaza la línea 'const VAR=...;' en el HTML.
    Maneja tanto JSON limpio ([...]) como JSON escapado (\[...\]) que
    puede generarse por ciertos editores o versiones previas del script."""
    # Patrón flexible: acepta backslash opcional antes de [ o {
    # Ejemplo: const ALL_DATA=\[\{...\}\]; o const ALL_DATA=[{...}];
    pattern = re.compile(
        r'(const\s+' + re.escape(var_name) + r'\s*=\s*)(\\?\[.*?\\?\]|\\?\{.*?\\?\})(;)',
        re.DOTALL
    )
    match = pattern.search(html)
    if not match:
        return html, False
    new_html = html[:match.start(2)] + new_json_str + html[match.end(2):]
    return new_html, True

def replace_fecha(html, nueva_fecha_str):
    """Reemplaza todas las ocurrencias de fecha DD/MM/YYYY en el dashboard"""
    old_pattern = re.compile(r'\d{2}/\d{2}/\d{4}')
    return old_pattern.sub(nueva_fecha_str, html)

def guardar(ruta, contenido):
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(contenido)

# ─────────────────────────────────────────────────────────────
#  STEP 1 – ENCONTRAR EL ARCHIVO EXCEL
# ─────────────────────────────────────────────────────────────
def encontrar_excel():
    # Buscar primero en carpeta data/ (modo GitHub), luego en la carpeta raíz (modo local)
    candidatos = []
    for carpeta in [CARPETA_DATA, CARPETA]:
        candidatos += [f for f in glob.glob(os.path.join(carpeta, "*.xlsx"))
                       if not os.path.basename(f).startswith("~$")]
    xlsx_files = list(dict.fromkeys(candidatos))  # dedup manteniendo orden
    if not xlsx_files:
        sys.exit("\n  ❌  No se encontró ningún archivo .xlsx en esta carpeta.\n"
                 "       Guardá el reporte SAP aquí y volvé a ejecutar el script.\n")

    if len(xlsx_files) == 1:
        return xlsx_files[0]

    # Si hay varios, usar automáticamente el más reciente
    xlsx_files.sort(key=os.path.getmtime, reverse=True)
    print(f"\n  Se encontraron {len(xlsx_files)} archivos Excel:")
    for i, f in enumerate(xlsx_files):
        mtime = datetime.fromtimestamp(os.path.getmtime(f)).strftime("%d/%m/%Y %H:%M")
        marca = "  ← SE USARÁ ESTE (más reciente)" if i == 0 else ""
        print(f"    [{i+1}] {os.path.basename(f)}  ({mtime}){marca}")
    print()
    return xlsx_files[0]   # Siempre el más reciente, sin preguntar

# ─────────────────────────────────────────────────────────────
#  STEP 2 – PARSEAR DATOS AP (Partidas Abiertas)
# ─────────────────────────────────────────────────────────────
def parsear_ap(wb):
    # Buscar hoja (nombre puede variar levemente)
    hoja_ap = None
    for name in wb.sheetnames:
        if "partida" in name.lower() or "abierta" in name.lower():
            hoja_ap = name
            break
    if not hoja_ap:
        # Intentar con la primera hoja como fallback
        hoja_ap = wb.sheetnames[0]
        log(f"⚠️  No se encontró 'Partidas Abiertas' — usando hoja '{hoja_ap}'")
    else:
        log(f"Hoja AP: '{hoja_ap}'")

    ws = wb[hoja_ap]
    all_rows = list(ws.iter_rows(values_only=True))
    if not all_rows:
        log("❌  Hoja AP vacía")
        return []

    # Leer encabezados
    raw_headers = all_rows[0]
    headers = [str(c).strip() if c else "" for c in raw_headers]
    log(f"Columnas AP ({len(headers)}): " +
        " | ".join(f"{i}:{h}" for i,h in enumerate(headers) if h)[:200])

    # Detectar columnas por nombre
    c_prov  = find_col(headers, ["ACREEDOR","PROVEEDOR","NOMBRE ACREEDOR","NOMBRE"])
    c_dias  = find_col(headers, ["ATRASAD","DÍAS ATRASAD","DIAS ATRASAD","DÍAS VENC","DIAS VENC","VENCIMIENTO","NETO","DÍAS","DIAS"])
    c_imp   = find_col(headers, ["IMPORTE EN M","IMPORTE M","IMPORTE"])
    c_mon   = find_col(headers, ["MONEDA","MON.","DIVISA"])
    c_fecha = find_col(headers, ["FECHA DE VENC","FECHA VENC","FECHA_VENC","VENC","FECHA CONT","FECHA DE CONT","F.CONT","FECHA_CONT","FECHA"])
    c_cta   = find_col(headers, ["CUENTA","POSICIÓN","POSICION","CLASE DE CUENTA"])
    c_cat   = find_col(headers, ["CATEGORÍA","CATEGORIA","CLASE","TIPO"])
    c_sub   = find_col(headers, ["SUBCATEGORÍA","SUBCATEGORIA","SUB"])

    # Fallback a índices originales si no se encontraron
    if c_prov  is None: c_prov  = 5
    if c_dias  is None: c_dias  = 6
    if c_imp   is None: c_imp   = 9
    if c_fecha is None: c_fecha = 13
    if c_cta   is None: c_cta   = 18
    if c_cat   is None: c_cat   = 19
    if c_sub   is None: c_sub   = 20

    log(f"  → Columnas detectadas: prov={c_prov} dias={c_dias} imp={c_imp} "
        f"mon={c_mon} fecha={c_fecha} cat={c_cat} sub={c_sub}")

    # Mostrar primera fila de datos para diagnóstico
    if len(all_rows) > 1:
        sample = all_rows[1]
        claves = {"prov":c_prov,"dias":c_dias,"imp":c_imp,"fecha":c_fecha,"cat":c_cat,"sub":c_sub}
        muestra = {k: (str(sample[v])[:25] if v is not None and v < len(sample) else "—")
                   for k, v in claves.items()}
        log(f"  → Muestra fila 1: {muestra}")

    records = []
    skipped = 0
    for row in all_rows[1:]:
        if not any(row):
            continue
        try:
            proveedor  = str(row[c_prov]  or "").strip() if c_prov  < len(row) else ""
            dias_raw   = row[c_dias]  if c_dias  < len(row) else 0
            importe_v  = row[c_imp]   if c_imp   < len(row) else None
            fecha_cont = row[c_fecha] if c_fecha < len(row) else None
            cuenta     = str(row[c_cta] or "").strip() if c_cta < len(row) else ""
            categoria  = str(row[c_cat] or "").strip().upper() if c_cat < len(row) else ""
            subcat     = str(row[c_sub] or "").strip() if c_sub < len(row) else ""
        except (IndexError, TypeError):
            skipped += 1
            continue

        if not proveedor:
            skipped += 1
            continue

        # Días: puede ser int, float o string
        try:
            dias = int(float(str(dias_raw or 0).replace(",",".")))
        except:
            dias = 0

        # Moneda: si hay columna separada, usarla
        moneda = "GS"
        if c_mon is not None and c_mon < len(row):
            mon_val = str(row[c_mon] or "").strip().upper()
            if "USD" in mon_val or mon_val in ("U","$","D"):
                moneda = "USD"

        # Importe
        if importe_v is None or importe_v == "":
            skipped += 1
            continue

        # Si hay columna de moneda separada, el importe es número crudo
        if c_mon is not None and isinstance(importe_v, (int, float)):
            monto = abs(float(importe_v))
        else:
            monto, moneda = parse_monto(importe_v)

        if monto == 0:
            skipped += 1
            continue

        records.append({
            "p":   proveedor,
            "c":   cuenta,
            "cat": categoria,
            "sub": subcat,
            "m":   monto,
            "mon": moneda,
            "d":   dias,
            "v":   dias > 0,
            "a":   aging_ap(dias),
            "f":   fecha_ym(fecha_cont)
        })

    log(f"AP: {len(records)} registros leídos  ({skipped} descartados)")
    if len(records) == 0:
        log("⚠️  ATENCIÓN: 0 registros AP. Verificá que el archivo Excel")
        log("    tenga la hoja correcta y las columnas esperadas.")
    return records

# ─────────────────────────────────────────────────────────────
#  STEP 3 – PARSEAR DATOS AR (Partidas abiertas Clientes)
# ─────────────────────────────────────────────────────────────
def parsear_ar(wb):
    from datetime import date as _date
    hoy = _date.today()

    # Buscar la hoja de clientes (puede variar el nombre exacto)
    hoja_cli = None
    for name in wb.sheetnames:
        if "cliente" in name.lower() or "cobrar" in name.lower() or "ar" == name.lower().strip():
            hoja_cli = name
            break
    if not hoja_cli:
        log("⚠️  No se encontró hoja de Clientes — se conserva la data existente.")
        return None

    log(f"Hoja AR: '{hoja_cli}'")
    ws = wb[hoja_cli]
    all_rows = list(ws.iter_rows(values_only=True))
    if not all_rows:
        return None

    raw_headers = all_rows[0]
    headers = [str(c).strip() if c else "" for c in raw_headers]
    log(f"Columnas AR ({len(headers)}): " +
        " | ".join(f"{i}:{h}" for i,h in enumerate(headers) if h)[:200])

    # Detectar columnas por nombre
    c_fecha  = find_col(headers, ["FECHA CONT","F.CONT","FECHA_CONT","FECHA CONTAB","FECHA DE CONTAB","FECHA"])
    c_fvenc  = find_col(headers, ["FECHA DE VENCIMIENTO","FECHA VENC","VENCIMIENTO","FECHA_VENC"])
    c_fac    = find_col(headers, ["FACTURA","N°FACTURA","NUMERO","DOCUMENTO","N.FACTURA","FAC"])
    c_cli    = find_col(headers, ["CLIENTE","DEUDOR","NOMBRE CLIENTE","NOMBRE"])
    c_saldo  = find_col(headers, ["SALDO","IMPORTE","MONTO","TOTAL"])
    c_dias   = find_col(headers, ["DÍAS ATRASO","DIAS ATRASO","DÍAS VENC","DIAS VENC","DÍAS","DIAS"])
    c_com    = find_col(headers, ["COMENTARIO","OBSERV","NOTA"])
    c_estado = find_col(headers, ["ESTADO","STATUS","SITUACION","SITUACIÓN"])
    c_cobro  = find_col(headers, ["FECHA COBRO","F.COBRO","COBRO","FECHA PAGO"])
    c_mes    = find_col(headers, ["MES"])
    c_anio   = find_col(headers, ["AÑO","ANIO","AÑO","YEAR"])

    # Fallback a índices originales
    if c_fecha  is None: c_fecha  = 0
    if c_fvenc  is None: c_fvenc  = 2
    if c_fac    is None: c_fac    = 3
    if c_cli    is None: c_cli    = 4
    if c_saldo  is None: c_saldo  = 5
    if c_dias   is None: c_dias   = 6
    if c_com    is None: c_com    = 7
    if c_estado is None: c_estado = 8
    if c_cobro  is None: c_cobro  = 9
    if c_mes    is None: c_mes    = 10
    if c_anio   is None: c_anio   = 11

    log(f"  → cli={c_cli} saldo={c_saldo} dias={c_dias} estado={c_estado} fac={c_fac}")

    records = []
    skipped = 0
    for row in all_rows[1:]:
        if not any(row):
            continue

        def get(col, default=""):
            try:
                v = row[col]
                return v if v is not None else default
            except IndexError:
                return default

        try:
            fecha_cont = get(c_fecha)
            fvenc_raw  = get(c_fvenc)
            factura    = str(get(c_fac, "")).strip()
            cliente    = str(get(c_cli, "")).strip()
            saldo_raw  = get(c_saldo, 0)
            dias_raw   = get(c_dias, 0)
            comentario = str(get(c_com, "")).strip().replace("-","").strip()
            estado     = str(get(c_estado, "")).strip().upper()
            cobro_raw  = get(c_cobro)
            mes_txt    = str(get(c_mes, "")).strip()
            anio_val   = get(c_anio)
        except Exception:
            skipped += 1
            continue

        if not cliente or not factura:
            skipped += 1
            continue

        try:
            saldo = abs(float(str(saldo_raw).replace(",","."))) if saldo_raw != "" else 0.0
        except:
            saldo = 0.0

        try:
            dias = int(float(str(dias_raw or 0).replace(",",".")))
        except:
            dias = 0

        # Normalizar estado
        if estado not in ("COBRADO","VENCIDO","A VENCER"):
            estado = "VENCIDO" if dias > 0 else "A VENCER"

        # Fecha de vencimiento real
        if isinstance(fvenc_raw, datetime):
            fvenc_str = fvenc_raw.date().strftime("%Y-%m-%d")
        elif isinstance(fvenc_raw, date):
            fvenc_str = fvenc_raw.strftime("%Y-%m-%d")
        elif isinstance(fvenc_raw, str) and len(fvenc_raw) >= 8:
            fvenc_str = fvenc_raw[:10]
        else:
            fvenc_str = ""

        # Corregir estado: SAP a veces deja "A VENCER" aunque la fecha ya pasó
        if estado == "A VENCER" and fvenc_str:
            try:
                fvenc_date = date.fromisoformat(fvenc_str)
                if fvenc_date < hoy:
                    estado = "VENCIDO"
                    if dias <= 0:
                        dias = (hoy - fvenc_date).days
            except Exception:
                pass

        # Fecha de cobro
        if isinstance(cobro_raw, (datetime, date)):
            cobro_str = cobro_raw.strftime("%Y-%m-%d")
        elif isinstance(cobro_raw, str) and len(cobro_raw) > 3:
            cobro_str = cobro_raw
        else:
            cobro_str = "PENDIENTE"

        # Año
        if isinstance(anio_val, (int, float)):
            anio_int = int(anio_val)
        elif isinstance(fecha_cont, (datetime, date)):
            anio_int = fecha_cont.year
        else:
            anio_int = 2026

        # Mes texto
        if not mes_txt and isinstance(fecha_cont, (datetime, date)):
            mes_txt = MESES.get(fecha_cont.month, "")

        incobrable = any(ex.upper() in cliente.upper() for ex in CLIENTES_INCOBRABLES)
        records.append({
            "cli":        cliente,
            "fac":        factura,
            "saldo":      round(saldo, 2),
            "dias":       dias,
            "estado":     estado,
            "aging":      aging_ar(dias, estado),
            "rating":     rating_ar(dias, estado),
            "f":          fecha_ym(fecha_cont),
            "fvenc":      fvenc_str,
            "mes":        mes_txt.lower(),
            "anio":       anio_int,
            "cobro":      cobro_str,
            "com":        comentario,
            "incobrable": incobrable
        })

    log(f"AR (Clientes): {len(records)} registros leídos  ({skipped} descartados)")
    if len(records) == 0:
        log("⚠️  ATENCIÓN: 0 registros AR. Verificá las columnas del Excel.")
    return records

# ─────────────────────────────────────────────────────────────
#  STEP 3b – PARSEAR CHEQUES DIFERIDOS
# ─────────────────────────────────────────────────────────────
def parsear_cheques(wb):
    """Parsea la hoja 'Cheques Diferidos'.
    Fila 1: encabezados, Datos desde fila 2.
    Columnas: # | Clave | Número | Importe | Nombre cuenta | Acreedor | Fecha Vto | Fecha Cont
    Retorna lista de {acreedor, numero, monto, moneda, fecha_vto}"""
    hoja_ch = None
    for name in wb.sheetnames:
        nl = name.lower()
        if "cheque" in nl or "diferido" in nl:
            hoja_ch = name
            break
    if not hoja_ch:
        log("⚠️  No se encontró hoja Cheques Diferidos — CH no se actualizará.")
        return []

    log(f"Hoja Cheques: '{hoja_ch}'")
    ws = wb[hoja_ch]
    all_rows = list(ws.iter_rows(values_only=True))
    if len(all_rows) < 2:
        log("⚠️  Hoja Cheques Diferidos vacía.")
        return []

    # Fila 0: encabezados, datos desde fila 1
    headers = [str(c).strip() if c else "" for c in all_rows[0]]
    log(f"Columnas CH ({len(headers)}): " +
        " | ".join(f"{i}:{h}" for i, h in enumerate(headers) if h))

    c_imp    = find_col(headers, ["IMPORTE","MONTO","TOTAL"])
    c_cuenta = find_col(headers, ["NOMBRE DE CUENTA","CUENTA","TIPO"])
    c_acr    = find_col(headers, ["ACREEDOR","PROVEEDOR","NOMBRE"])
    c_vto    = find_col(headers, ["VENCIMIENTO","FECHA VTO","FECHA_VTO"])
    c_num    = find_col(headers, ["NÚMERO","NUMERO","NRO","N°","CHEQUE"])

    if c_imp    is None: c_imp    = 3
    if c_cuenta is None: c_cuenta = 4
    if c_acr    is None: c_acr    = 5
    if c_vto    is None: c_vto    = 6
    if c_num    is None: c_num    = 2

    records = []
    skipped = 0
    for row in all_rows[1:]:
        if not any(row):
            continue
        try:
            monto_raw = row[c_imp]    if c_imp    < len(row) else None
            cuenta    = str(row[c_cuenta] or "").strip() if c_cuenta < len(row) else ""
            acreedor  = str(row[c_acr]    or "").strip() if c_acr    < len(row) else ""
            vto_raw   = row[c_vto]    if c_vto    < len(row) else None
            numero    = str(row[c_num]    or "").strip() if c_num    < len(row) else ""
        except (IndexError, TypeError):
            skipped += 1
            continue

        # Solo cheques diferidos (GS o USD)
        cuenta_up = cuenta.upper()
        if "CHEQUES DIFERIDOS" not in cuenta_up and "CHEQUE DIFERIDO" not in cuenta_up:
            skipped += 1
            continue

        moneda = "USD" if "USD" in cuenta_up else "GS"

        # Monto
        if monto_raw is None or monto_raw == "":
            skipped += 1
            continue
        try:
            monto = abs(float(str(monto_raw).replace(",", ".")))
        except Exception:
            monto, _ = parse_monto(monto_raw)

        if monto == 0:
            skipped += 1
            continue

        # Fecha de vencimiento
        if isinstance(vto_raw, datetime):
            fecha_dt = vto_raw.date()
        elif isinstance(vto_raw, date):
            fecha_dt = vto_raw
        elif isinstance(vto_raw, str) and len(vto_raw) > 3:
            try:
                fecha_dt = datetime.strptime(vto_raw[:10], "%Y-%m-%d").date()
            except Exception:
                skipped += 1
                continue
        else:
            skipped += 1
            continue

        records.append({
            "acreedor": acreedor,
            "numero":   numero,
            "monto":    round(monto, 2),
            "moneda":   moneda,
            "fecha":    fecha_dt.strftime("%Y-%m-%d")
        })

    log(f"CH: {len(records)} cheques diferidos leídos  ({skipped} descartados)")
    return records


# ─────────────────────────────────────────────────────────────
#  STEP 3c – PARSEAR MATERIA PRIMA
# ─────────────────────────────────────────────────────────────
def parsear_mp(wb):
    """Parsea la hoja 'Materia Prima'.
    Fila 1: título (skip), Fila 2: encabezados, Datos desde fila 3.
    Retorna lista de {fecha, proveedor, po, categoria, monto}"""
    hoja_mp = None
    for name in wb.sheetnames:
        nl = name.lower()
        if "materia" in nl or "prima" in nl or nl.strip() == "mp":
            hoja_mp = name
            break
    if not hoja_mp:
        log("⚠️  No se encontró hoja Materia Prima — MP no se actualizará.")
        return []

    log(f"Hoja MP: '{hoja_mp}'")
    ws = wb[hoja_mp]
    all_rows = list(ws.iter_rows(values_only=True))
    if len(all_rows) < 3:
        log("⚠️  Hoja Materia Prima con menos de 3 filas — sin datos.")
        return []

    # Fila 0: título, Fila 1: encabezados
    headers = [str(c).strip() if c else "" for c in all_rows[1]]
    log(f"Columnas MP ({len(headers)}): " +
        " | ".join(f"{i}:{h}" for i, h in enumerate(headers) if h))

    c_fecha = find_col(headers, ["FECHA DE PAGO", "FECHA PAGO", "FECHA"])
    c_prov  = find_col(headers, ["PROVEEDOR", "ACREEDOR", "NOMBRE"])
    c_po    = find_col(headers, ["ORDEN", "PO", "N°"])
    c_cat   = find_col(headers, ["CATEGORÍA", "CATEGORIA", "TIPO", "CLASE"])
    c_monto = find_col(headers, ["MONTO", "IMPORTE", "TOTAL"])

    if c_fecha is None: c_fecha = 0
    if c_prov  is None: c_prov  = 1
    if c_po    is None: c_po    = 2
    if c_cat   is None: c_cat   = 3
    if c_monto is None: c_monto = 4

    records = []
    skipped = 0
    for row in all_rows[2:]:
        if not any(row):
            continue
        try:
            fecha_raw = row[c_fecha] if c_fecha < len(row) else None
            proveedor = str(row[c_prov] or "").strip() if c_prov < len(row) else ""
            po        = str(row[c_po]   or "").strip() if c_po   < len(row) else ""
            categoria = str(row[c_cat]  or "").strip() if c_cat  < len(row) else ""
            monto_raw = row[c_monto] if c_monto < len(row) else None
        except (IndexError, TypeError):
            skipped += 1
            continue

        # Fecha de pago
        if isinstance(fecha_raw, datetime):
            fecha_dt = fecha_raw.date()
        elif isinstance(fecha_raw, date):
            fecha_dt = fecha_raw
        elif isinstance(fecha_raw, str) and len(fecha_raw) > 3:
            try:
                fecha_dt = datetime.strptime(fecha_raw[:10], "%Y-%m-%d").date()
            except Exception:
                skipped += 1
                continue
        else:
            skipped += 1
            continue

        # Monto
        if monto_raw is None or monto_raw == "":
            skipped += 1
            continue
        try:
            monto = abs(float(str(monto_raw).replace(",", ".")))
        except Exception:
            monto, _ = parse_monto(monto_raw)

        if monto == 0 or not proveedor:
            skipped += 1
            continue

        records.append({
            "fecha":     fecha_dt.strftime("%Y-%m-%d"),
            "proveedor": proveedor,
            "po":        po,
            "categoria": categoria,
            "monto":     round(monto, 2)
        })

    log(f"MP: {len(records)} pagos leídos  ({skipped} descartados)")
    return records


# ─────────────────────────────────────────────────────────────
#  STEP 3c – PARSEAR GASTOS FIJOS
# ─────────────────────────────────────────────────────────────
def parsear_gf(wb):
    """Parsea la hoja 'Gastos Fijos'.
    Fila 1: título (skip), Fila 2: instrucción (skip), Fila 3: encabezados,
    Datos desde fila 4. Solo incluye filas con Activo='Sí'.
    Retorna lista de {concepto, clasif, monto, frecuencia, regla, mes_inicio, activo}"""
    hoja_gf = None
    for name in wb.sheetnames:
        nl = name.lower()
        if "gasto" in nl or "fijo" in nl or nl.strip() == "gf":
            hoja_gf = name
            break
    if not hoja_gf:
        log("⚠️  No se encontró hoja Gastos Fijos — GF no se actualizará.")
        return []

    log(f"Hoja GF: '{hoja_gf}'")
    ws = wb[hoja_gf]
    all_rows = list(ws.iter_rows(values_only=True))
    if len(all_rows) < 4:
        log("⚠️  Hoja Gastos Fijos con menos de 4 filas — sin datos.")
        return []

    # Fila 0: título, Fila 1: instrucción, Fila 2: encabezados
    headers = [str(c).strip() if c else "" for c in all_rows[2]]
    log(f"Columnas GF ({len(headers)}): " +
        " | ".join(f"{i}:{h}" for i, h in enumerate(headers) if h))

    c_conc   = find_col(headers, ["PROVEEDOR", "CONCEPTO", "NOMBRE"])
    c_clasif = find_col(headers, ["CLASIFICACIÓN", "CLASIFICACION", "CLASE", "TIPO"])
    c_monto  = find_col(headers, ["MONTO", "IMPORTE", "TOTAL"])
    c_frec   = find_col(headers, ["FRECUENCIA", "FREQ"])
    c_regla  = find_col(headers, ["REGLA", "PAGO", "VENCIMIENTO"])
    c_mes    = find_col(headers, ["MES DE INICIO", "MES INICIO", "MES"])
    c_activo = find_col(headers, ["ACTIVO", "ACTIVE", "ESTADO"])

    if c_conc   is None: c_conc   = 0
    if c_clasif is None: c_clasif = 1
    if c_monto  is None: c_monto  = 2
    if c_frec   is None: c_frec   = 3
    if c_regla  is None: c_regla  = 4
    if c_mes    is None: c_mes    = 5
    if c_activo is None: c_activo = 6

    records = []
    skipped = 0
    for row in all_rows[3:]:
        if not any(row):
            continue
        try:
            concepto  = str(row[c_conc]   or "").strip() if c_conc   < len(row) else ""
            clasif    = str(row[c_clasif] or "").strip() if c_clasif < len(row) else ""
            monto_raw = row[c_monto] if c_monto < len(row) else None
            frecuencia= str(row[c_frec]   or "Mensual").strip() if c_frec < len(row) else "Mensual"
            regla     = str(row[c_regla]  or "").strip() if c_regla  < len(row) else ""
            mes_ini   = str(row[c_mes]    or "").strip() if c_mes    < len(row) else ""
            activo_raw= str(row[c_activo] or "").strip() if c_activo < len(row) else ""
        except (IndexError, TypeError):
            skipped += 1
            continue

        # Saltar filas sin concepto real (TOTAL, instrucciones, etc.)
        if not concepto or concepto.upper() in ("TOTAL", "TOTALES", "NOTA", "INSTRUCCIÓN"):
            skipped += 1
            continue

        # Solo activos
        if activo_raw.lower() not in ("sí", "si", "yes", "1", "true", "activo", "s"):
            skipped += 1
            continue

        # Monto
        if monto_raw is None or monto_raw == "":
            skipped += 1
            continue
        try:
            monto = abs(float(str(monto_raw).replace(",", ".")))
        except Exception:
            monto, _ = parse_monto(monto_raw)

        if monto == 0:
            skipped += 1
            continue

        records.append({
            "concepto":   concepto,
            "clasif":     clasif,
            "monto":      round(monto, 2),
            "frecuencia": frecuencia,
            "regla":      regla,
            "mes_inicio": mes_ini,
            "activo":     True
        })

    log(f"GF: {len(records)} gastos fijos activos leídos  ({skipped} descartados)")
    return records


# ─────────────────────────────────────────────────────────────
#  STEP 4 – CALCULAR RESUMEN AGREGADO
# ─────────────────────────────────────────────────────────────
def calcular_resumen(ap_records, ar_records):
    # AP (sin FINANCIAMIENTO)
    ap = [r for r in ap_records if r["cat"] != "FINANCIAMIENTO"]
    fin= [r for r in ap_records if r["cat"] == "FINANCIAMIENTO"]

    def sumas(lst):
        gs = sum(r["m"] for r in lst if r["mon"]=="GS")
        usd= sum(r["m"] for r in lst if r["mon"]=="USD")
        n  = len(lst)
        vgs= sum(r["m"] for r in lst if r["mon"]=="GS" and r["v"])
        vu = sum(r["m"] for r in lst if r["mon"]=="USD" and r["v"])
        nv = sum(1 for r in lst if r["v"])
        xvgs= sum(r["m"] for r in lst if r["mon"]=="GS" and not r["v"])
        xvu = sum(r["m"] for r in lst if r["mon"]=="USD" and not r["v"])
        ag_gs = {b: sum(r["m"] for r in lst if r["mon"]=="GS" and r["a"]==b)
                 for b in AGING_BUCKETS}
        ag_usd= {b: sum(r["m"] for r in lst if r["mon"]=="USD" and r["a"]==b)
                 for b in AGING_BUCKETS}
        return {"gs":gs,"usd":usd,"n":n,"venc_gs":vgs,"venc_usd":vu,
                "n_venc":nv,"xvenc_gs":xvgs,"xvenc_usd":xvu,
                "aging_gs":ag_gs,"aging_usd":ag_usd}

    ap_sum  = sumas(ap)
    fin_sum = sumas(fin)

    # Bancos breakdown (FIN)
    bancos = {}
    for r in fin:
        k = r["p"]
        bancos[k] = bancos.get(k, 0) + (r["m"] if r["mon"]=="USD" else 0)
    fin_sum["bancos"] = bancos

    # AR
    total_ar = sum(r["saldo"] for r in ar_records)
    cobrado  = sum(r["saldo"] for r in ar_records if r["estado"]=="COBRADO")
    vencido  = sum(r["saldo"] for r in ar_records if r["estado"]=="VENCIDO")
    avencer  = sum(r["saldo"] for r in ar_records if r["estado"]=="A VENCER")
    n_ar     = len(ar_records)
    n_venc   = sum(1 for r in ar_records if r["estado"]=="VENCIDO")

    ar_aging = {}
    for r in ar_records:
        k = r["aging"]
        ar_aging[k] = ar_aging.get(k,0) + r["saldo"]

    # Top 5 vencidos
    venc_cli = {}
    for r in ar_records:
        if r["estado"] == "VENCIDO":
            venc_cli[r["cli"]] = venc_cli.get(r["cli"],0) + r["saldo"]
    top_venc = sorted(venc_cli.items(), key=lambda x: -x[1])[:5]

    # Top 5 total
    total_cli = {}
    for r in ar_records:
        total_cli[r["cli"]] = total_cli.get(r["cli"],0) + r["saldo"]
    top_total = sorted(total_cli.items(), key=lambda x: -x[1])[:5]

    # Promedio de atraso — incluye VENCIDO (mora actual) + COBRADO (mora histórica)
    # Excluye CAPITAL TRADE por ser caso atípico
    EXCLUIR_PROMEDIO = ["CAPITAL TRADE"]
    dias_todos = [r["dias"] for r in ar_records
                  if r["estado"] in ("VENCIDO", "COBRADO") and r["dias"] > 0
                  and not any(ex in r["cli"].upper() for ex in EXCLUIR_PROMEDIO)]
    avg_delay = round(sum(dias_todos) / len(dias_todos)) if dias_todos else 0

    # Desglose: solo vencidos actuales
    dias_venc_solo = [r["dias"] for r in ar_records
                      if r["estado"] == "VENCIDO" and r["dias"] > 0
                      and not any(ex in r["cli"].upper() for ex in EXCLUIR_PROMEDIO)]
    avg_delay_vencido = round(sum(dias_venc_solo) / len(dias_venc_solo)) if dias_venc_solo else 0

    # Solo cobrados históricos
    dias_cob_solo = [r["dias"] for r in ar_records
                     if r["estado"] == "COBRADO" and r["dias"] > 0
                     and not any(ex in r["cli"].upper() for ex in EXCLUIR_PROMEDIO)]
    avg_delay_cobrado = round(sum(dias_cob_solo) / len(dias_cob_solo)) if dias_cob_solo else 0

    # ── DSO: días ponderados por saldo (vencido + cobrado, excluye Capital Trade) ──
    dso_items = [(r["saldo"], r["dias"]) for r in ar_records
                 if r["estado"] in ("VENCIDO", "COBRADO") and r["dias"] > 0
                 and not any(ex in r["cli"].upper() for ex in EXCLUIR_PROMEDIO)]
    dso_num = sum(s * d for s, d in dso_items)
    dso_den = sum(s for s, d in dso_items)
    dso = round(dso_num / dso_den) if dso_den > 0 else 0

    ar_sum = {
        "total":total_ar,"cobrado":cobrado,"vencido":vencido,
        "avencer":avencer,"n":n_ar,"n_venc":n_venc,"aging":ar_aging,
        "avg_delay":avg_delay,
        "avg_delay_vencido":avg_delay_vencido,
        "avg_delay_cobrado":avg_delay_cobrado,
        "dso":dso,
        "top_venc": [{"cli":c,"s":s} for c,s in top_venc],
        "top_total":[{"cli":c,"s":s} for c,s in top_total]
    }

    # ── DPO: días ponderados por monto (AP operativa vencida) ──
    def to_usd(r, tc=TC_REF):
        return r["m"] if r["mon"]=="USD" else r["m"]/tc
    dpo_items = [(to_usd(r), r["d"]) for r in ap if r["d"] > 0]
    dpo_num = sum(m * d for m, d in dpo_items)
    dpo_den = sum(m for m, d in dpo_items)
    dpo = round(dpo_num / dpo_den) if dpo_den > 0 else 0
    ap_sum["dpo"] = dpo

    return {"ap": ap_sum, "fin": fin_sum, "ar": ar_sum}


# ─────────────────────────────────────────────────────────────
#  STEP 4b – ACTUALIZAR HISTORIAL DE DEUDA
# ─────────────────────────────────────────────────────────────
def actualizar_historial(resumen, fecha_str):
    """Agrega el snapshot de hoy al historial y lo devuelve para inyectar en el HTML."""
    ap  = resumen["ap"]
    fin = resumen["fin"]
    ar  = resumen["ar"]

    ap_usd  = round(ap["usd"]  + ap["gs"]  / TC_REF, 2)
    fin_usd = round(fin["usd"] + fin["gs"] / TC_REF, 2)
    ar_pend = round(ar["vencido"] + ar["avencer"], 2)
    net     = round(ar_pend - ap_usd - fin_usd, 2)

    nuevo = {"f": fecha_str, "ap": ap_usd, "fin": fin_usd, "ar": ar_pend, "net": net}

    # Cargar historial existente
    historial = []
    if os.path.exists(HISTORIAL_PATH):
        try:
            with open(HISTORIAL_PATH, "r", encoding="utf-8") as f:
                historial = json.load(f)
        except Exception:
            historial = []

    # Evitar duplicar la misma fecha
    historial = [h for h in historial if h.get("f") != fecha_str]
    historial.append(nuevo)

    # Guardar (solo los últimos 104 puntos ≈ 2 años de datos semanales)
    historial = historial[-104:]
    with open(HISTORIAL_PATH, "w", encoding="utf-8") as f:
        json.dump(historial, f, ensure_ascii=False, separators=(",", ":"))

    log(f"📈 Historial actualizado: {len(historial)} punto(s) guardado(s)")
    return historial

# ─────────────────────────────────────────────────────────────
#  STEP 5 – INYECTAR EN LOS HTML
# ─────────────────────────────────────────────────────────────
def actualizar_html(ruta, var_name, nueva_data, nueva_fecha, descripcion):
    if not os.path.exists(ruta):
        log(f"⚠️  No se encontró {os.path.basename(ruta)} — saltando.")
        return

    with open(ruta, "r", encoding="utf-8") as f:
        html = f.read()

    # Actualizar fecha PRIMERO (antes de inyectar datos) para no tocar
    # las fechas dentro del JSON del historial que se inyectará a continuación
    html = replace_fecha(html, nueva_fecha)

    # Reemplazar bloque de datos
    nuevo_json = json.dumps(nueva_data, ensure_ascii=False, separators=(',',':'))
    html_nuevo, ok = replace_data_block(html, var_name, nuevo_json)

    if not ok:
        log(f"⚠️  No se pudo reemplazar {var_name} en {os.path.basename(ruta)}")
        return

    guardar(ruta, html_nuevo)
    log(f"✅  {descripcion} actualizado  ({len(nueva_data) if isinstance(nueva_data, list) else 'OK'} registros)")

# ─────────────────────────────────────────────────────────────
#  STEP 3d – PARSEAR FLUJO MP (nueva hoja con egresos detallados)
# ─────────────────────────────────────────────────────────────
def parsear_flujo_mp(wb):
    """Lee la hoja 'FLUJO MP' y extrae todos los egresos con fecha y categoría.
    Retorna lista de dicts: {fecha, desc, monto, cat}
      cat: 'mp' = proveedores/MP  |  'gf' = nómina/financiero  |  'ch' = cheques diferidos
    Retorna None si la hoja no existe (fallback al método tradicional).
    """
    SHEET_NAMES = ["FLUJO MP", "Flujo MP", "FLUJO_MP", "flujo mp"]
    ws = None
    for name in SHEET_NAMES:
        if name in wb.sheetnames:
            ws = wb[name]
            break
    if ws is None:
        log("ℹ️  Hoja 'FLUJO MP' no encontrada — se usarán las hojas Materia Prima y Gastos Fijos.")
        return None

    records = []
    skipped = 0

    # Palabras clave para nómina/financiero
    KW_GF = ["SALARIO", "QUINCENA", "IPS", "NOMINA",
              "INTERESES", "INTERES", "PRESTAMO", "PRÉSTAMO", "SOLAR", "ITAU", "BANCOP"]

    for row in ws.iter_rows(min_row=5, values_only=True):
        if len(row) < 7:
            continue
        fecha_raw, nsem, cod, desc, ing, eg, saldo = row[0], row[1], row[2], row[3], row[4], row[5], row[6]

        # Necesita fecha válida
        if not isinstance(fecha_raw, (datetime, date)):
            continue
        fecha_dt = fecha_raw.date() if isinstance(fecha_raw, datetime) else fecha_raw

        desc_str = str(desc).strip() if desc else ""
        desc_up  = desc_str.upper()

        # Saltar cobros de clientes (ya manejados por AR)
        if "⬆ COBRO" in desc_up or desc_up.startswith("COBRO:"):
            continue
        # Saltar subtotales y encabezados de mes
        if "SUBTOTAL" in desc_up or (not eg and not ing):
            continue

        # Solo procesar egresos
        try:
            eg_val = float(eg) if eg else 0.0
        except Exception:
            eg_val = 0.0
        if eg_val <= 0:
            skipped += 1
            continue

        # Categorizar
        cod_str = str(cod).upper().strip() if cod else ""
        if "CHEQUE DIF" in cod_str:
            cat = "ch"
        elif any(kw in desc_up for kw in KW_GF):
            cat = "gf"
        else:
            cat = "mp"

        records.append({
            "fecha": fecha_dt,
            "desc":  desc_str,
            "cod":   cod_str,
            "monto": round(eg_val, 2),
            "cat":   cat
        })

    log(f"FLUJO MP: {len(records)} egresos leídos  ({skipped} sin monto)")
    return records

# ─────────────────────────────────────────────────────────────
#  FLUJO DE CAJA — actualización de AR, MP y GF
# ─────────────────────────────────────────────────────────────
def actualizar_flujo(ar_records, mp_records, gf_records, ch_records, fecha_str, flujo_mp_records=None):
    """Actualiza cobros AR, Materia Prima, Gastos Fijos y Cheques Diferidos en Dashboard_Flujo.html."""
    if not os.path.exists(DASH_FLUJO):
        log(f"⚠️  No se encontró Dashboard_Flujo.html — saltando.")
        return

    import calendar as _cal
    from datetime import date, timedelta
    hoy = date.today()

    # 13 semanas comenzando el lunes de esta semana
    lunes = hoy - timedelta(days=hoy.weekday())
    WEEK_STARTS = [lunes + timedelta(weeks=w) for w in range(13)]
    n_weeks = len(WEEK_STARTS)

    # ── Labels ────────────────────────────────────────────────
    MES_ABR = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
    labels = []
    for w, ws in enumerate(WEEK_STARTS):
        iso_week = ws.isocalendar()[1]
        labels.append(f"Sem {iso_week} ({ws.strftime('%d/')}{MES_ABR[ws.month-1]})")

    def find_week(d):
        """Retorna el índice de semana (0-12) para una fecha, o None si fuera de rango."""
        if d < WEEK_STARTS[0]:
            return None
        for w in range(n_weeks - 1):
            if d < WEEK_STARTS[w + 1]:
                return w
        # Dentro de la última semana
        if d <= WEEK_STARTS[-1] + timedelta(days=6):
            return n_weeks - 1
        return None

    # ── AR ────────────────────────────────────────────────────
    ar_tot   = [0.0] * n_weeks
    ar_items = {str(w): [] for w in range(n_weeks)}

    for r in ar_records:
        estado = r.get("estado", "")
        if estado == "COBRADO":
            continue  # ya cobrado → no hay flujo pendiente

        # Excluir clientes incobrables (quiebra, deuda irrecuperable)
        cli_upper = r.get("cli", "").upper()
        if any(ex.upper() in cli_upper for ex in CLIENTES_INCOBRABLES):
            continue

        dias_val = r.get("dias", 0)
        if dias_val is None:
            continue
        try:
            dias_int = int(dias_val)
        except Exception:
            continue

        # VENCIDO (dias > 0) → ya pasó la fecha, acumular en semana actual
        # A VENCER (dias < 0) → fecha futura, asignar por vencimiento
        if estado == "VENCIDO" or dias_int > 0:
            # Ya vencido → acumular en semana actual pendiente de cobro
            due = hoy
            overdue = True
        else:
            # A vencer → usar FECHA DE VENCIMIENTO real del Excel
            fvenc_str = r.get("fvenc", "")
            if fvenc_str:
                try:
                    due = date.fromisoformat(fvenc_str)
                except Exception:
                    due = hoy + timedelta(days=abs(dias_int))
            elif dias_int < 0:
                # Tiene dias negativos pero sin fvenc → calcular desde dias
                due = hoy + timedelta(days=abs(dias_int))
            else:
                # Sin fvenc y sin dias → EN STOCK / EN PRODUCCIÓN, sin fecha definida
                continue  # excluir del flujo
            overdue = False

        w = find_week(due)
        if w is None:
            continue  # fuera del horizonte de 13 semanas → no incluir

        saldo = r.get("saldo", 0) or 0
        ar_tot[w] += saldo
        ar_items[str(w)].append({
            "cli":     r.get("cli", ""),
            "saldo":   round(float(saldo), 2),
            "due":     r.get("fvenc", due.strftime("%Y-%m-%d")),
            "overdue": overdue
        })

    ar_tot = [round(v, 2) for v in ar_tot]

    # ── MP / GF / CH ─────────────────────────────────────────
    mp_tot   = [0.0] * n_weeks
    mp_items = {str(w): [] for w in range(n_weeks)}
    gf_tot   = [0.0] * n_weeks
    gf_items = {str(w): [] for w in range(n_weeks)}
    ch_tot   = [0.0] * n_weeks
    ch_items = {str(w): [] for w in range(n_weeks)}
    total_gf_mes = 0.0

    if flujo_mp_records is not None:
        # ── FLUJO MP: fuente principal de egresos ────────────
        log("  Usando hoja FLUJO MP para egresos (MP + GF + CH)")
        for r in flujo_mp_records:
            w = find_week(r["fecha"])
            if w is None:
                continue  # fuera del horizonte de 13 semanas
            monto = r["monto"]
            item  = {"prov": r["desc"], "monto": monto, "fecha": r["fecha"].strftime("%Y-%m-%d")}
            cat   = r["cat"]
            if cat == "mp":
                mp_tot[w] += monto
                mp_items[str(w)].append(item)
            elif cat == "gf":
                gf_tot[w] += monto
                gf_items[str(w)].append(item)
                total_gf_mes += monto
            elif cat == "ch":
                ch_tot[w] += monto
                ch_items[str(w)].append(item)

    else:
        # ── FALLBACK: hojas Materia Prima + Gastos Fijos + Cheques ──
        log("  Usando hojas Materia Prima / Gastos Fijos (FLUJO MP no disponible)")

        for r in mp_records:
            try:
                fecha_mp = date.fromisoformat(r["fecha"])
            except Exception:
                continue
            w = find_week(fecha_mp)
            if w is None:
                continue
            monto = r.get("monto", 0) or 0
            mp_tot[w] += monto
            mp_items[str(w)].append({
                "prov":  r.get("proveedor", ""),
                "monto": round(float(monto), 2),
                "fecha": r["fecha"]
            })

        total_gf_mes = round(sum(g.get("monto", 0) for g in gf_records), 2)
        if gf_records:
            window_start = WEEK_STARTS[0]
            window_end   = WEEK_STARTS[-1] + timedelta(days=6)
            months_in_window = []
            d = window_start.replace(day=1)
            while d <= window_end:
                months_in_window.append((d.year, d.month))
                mo_next = d.month + 1 if d.month < 12 else 1
                yr_next = d.year + (1 if d.month == 12 else 0)
                d = d.replace(year=yr_next, month=mo_next)
            for yr, mo in months_in_window:
                last_day = _cal.monthrange(yr, mo)[1]
                for gf in gf_records:
                    regla = gf.get("regla", "").lower()
                    monto = gf.get("monto", 0) or 0
                    if "primera" in regla or "primer" in regla:
                        pay_date = date(yr, mo, 3)
                    elif "último" in regla or "ultimo" in regla:
                        pay_date = date(yr, mo, last_day)
                    elif "14" in regla:
                        pay_date = date(yr, mo, 14)
                    elif "23" in regla:
                        pay_date = date(yr, mo, 23)
                    else:
                        pay_date = date(yr, mo, 1)
                    if pay_date < window_start or pay_date > window_end:
                        continue
                    w = find_week(pay_date)
                    if w is None:
                        continue
                    gf_tot[w] += monto
                    gf_items[str(w)].append({
                        "concepto": gf.get("concepto", ""),
                        "monto":    round(float(monto), 2),
                        "fecha":    pay_date.strftime("%Y-%m-%d"),
                        "regla":    gf.get("regla", "")
                    })

        for r in ch_records:
            try:
                fecha_ch = date.fromisoformat(r["fecha"])
            except Exception:
                continue
            w = find_week(fecha_ch)
            if w is None:
                continue
            moneda    = r.get("moneda", "GS")
            monto_raw = r.get("monto", 0) or 0
            monto_usd = round(monto_raw / TC_REF, 2) if moneda == "GS" else round(float(monto_raw), 2)
            ch_tot[w] += monto_usd
            ch_items[str(w)].append({
                "acreedor": r.get("acreedor", ""),
                "numero":   r.get("numero", ""),
                "monto":    monto_usd,
                "moneda":   moneda,
                "fecha":    r["fecha"]
            })

    mp_tot = [round(v, 2) for v in mp_tot]
    gf_tot = [round(v, 2) for v in gf_tot]
    ch_tot = [round(v, 2) for v in ch_tot]

    # ── Leer DATA actual del HTML ──────────────────────────────
    with open(DASH_FLUJO, "r", encoding="utf-8") as f:
        html = f.read()

    m = re.search(r'const DATA\s*=\s*(\{.*?\});', html, re.DOTALL)
    if not m:
        log("⚠️  No se encontró const DATA en Dashboard_Flujo.html")
        return

    import json as _json
    try:
        data = _json.loads(m.group(1))
    except Exception as e:
        log(f"⚠️  Error al parsear DATA en Flujo: {e}")
        return

    # ── Inyectar nuevos valores ────────────────────────────────
    data["labels"]      = labels
    data["ar"]          = ar_tot
    data["arItems"]     = ar_items
    data["mp"]          = mp_tot
    data["mpItems"]     = mp_items
    data["gf"]          = gf_tot
    data["gfItems"]     = gf_items
    data["ch"]          = ch_tot
    data["chItems"]     = ch_items
    data["gastosFijos"] = [round(g.get("monto", 0), 2) for g in gf_records]
    data["totalGFMes"]  = total_gf_mes

    # Recalcular saldo acumulado semana a semana (AR - MP - GF - CH)
    si = data.get("saldoInicial", 0)
    saldos = []
    s = si
    for w in range(n_weeks):
        s += ar_tot[w] - mp_tot[w] - gf_tot[w] - ch_tot[w]
        saldos.append(round(s, 2))
    data["saldo"] = saldos
    data["s30"]   = saldos[min(3,  n_weeks-1)]
    data["s60"]   = saldos[min(7,  n_weeks-1)]
    data["s90"]   = saldos[min(11, n_weeks-1)]

    nuevo_json = _json.dumps(data, ensure_ascii=False, separators=(',', ':'))
    html_nuevo, ok = replace_data_block(html, "DATA", nuevo_json)
    if not ok:
        log("⚠️  No se pudo reemplazar DATA en Dashboard_Flujo.html")
        return

    html_nuevo = replace_fecha(html_nuevo, fecha_str)
    guardar(DASH_FLUJO, html_nuevo)
    cobros_av  = sum(1 for r in ar_records if r.get("estado") == "A VENCER")
    log(f"✅  Flujo de Caja actualizado  (AR:{cobros_av} cobros | MP:{len(mp_records)} pagos | GF:{len(gf_records)} GF | CH:{len(ch_records)} cheques)")

# ─────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────
def main():
    print("\n" + "═"*56)
    print("  ACTUALIZADOR DASHBOARD  —  CMA Paraguay S.A.")
    print("═"*56)

    # 1. Instalar openpyxl si falta
    try:
        import openpyxl
    except ImportError:
        print("\n  Instalando openpyxl...")
        os.system(f'"{sys.executable}" -m pip install openpyxl -q')
        import openpyxl

    # 2. Encontrar Excel
    excel_path = encontrar_excel()
    log(f"📂 Archivo: {os.path.basename(excel_path)}")

    # 3. Cargar workbook
    log("Leyendo Excel con SAP...")
    try:
        wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    except Exception as e:
        sys.exit(f"\n  ❌  No se pudo abrir el archivo: {e}\n"
                 f"       Asegurate de que no esté abierto en Excel.\n")

    log(f"Hojas encontradas: {', '.join(wb.sheetnames)}")

    # 4. Parsear datos
    print()
    ap_records      = parsear_ap(wb)
    ar_records      = parsear_ar(wb) or []
    ch_records      = parsear_cheques(wb)
    mp_records      = parsear_mp(wb)
    gf_records      = parsear_gf(wb)
    flujo_mp_records = parsear_flujo_mp(wb)   # Nueva hoja con egresos detallados

    # 5. Calcular resumen
    resumen = calcular_resumen(ap_records, ar_records)

    # 6. Fecha de hoy para los dashboards
    hoy = datetime.today()
    fecha_str = hoy.strftime("%d/%m/%Y")

    # 6b. Actualizar historial y adjuntarlo al resumen
    print()
    historial = actualizar_historial(resumen, fecha_str)
    resumen["historial"] = historial

    # 7. Actualizar los 4 HTMLs
    print()
    log("Actualizando dashboards...")

    # Operaciones (excluye FINANCIAMIENTO)
    actualizar_html(DASH_OPS,      "ALL_DATA",
                    ap_records,    fecha_str, "Dashboard_Operaciones")

    # Bancos (todos los registros, el HTML filtra cat===FINANCIAMIENTO)
    actualizar_html(DASH_BANCOS,   "ALL_DATA",
                    ap_records,    fecha_str, "Dashboard_Bancos")

    # Clientes
    if ar_records:
        actualizar_html(DASH_CLIENTES, "ALL_DATA",
                        ar_records,    fecha_str, "Dashboard_Clientes")
    else:
        log("⚠️  Clientes: sin datos, conservando dashboard existente.")

    # Flujo de Caja (AR + FLUJO MP o MP/GF/CH tradicional)
    if ar_records or flujo_mp_records or mp_records or gf_records or ch_records:
        actualizar_flujo(ar_records, mp_records, gf_records, ch_records, fecha_str,
                         flujo_mp_records=flujo_mp_records)
    else:
        log("⚠️  Flujo de Caja: sin datos, conservando dashboard existente.")

    # Resumen
    actualizar_html(DASH_RESUMEN,  "D",
                    resumen,       fecha_str, "Dashboard_Resumen")

    print()
    print("═"*56)
    print(f"  ✅  5 dashboards actualizados al {fecha_str}")
    print(f"  AP: {len(ap_records)} registros  |  AR: {len(ar_records)} registros")
    print(f"  MP: {len(mp_records)} pagos  |  GF: {len(gf_records)} gastos fijos  |  CH: {len(ch_records)} cheques dif.")
    print("  Resumen · C. a Pagar · Bancos · Clientes · Flujo de Caja")
    print("═"*56)
    print()
    print("  Dashboards listos y publicados en GitHub Pages.")
    print()

if __name__ == "__main__":
    main()
