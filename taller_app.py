import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import json
from collections import Counter
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Supabase
try:
    from supabase import create_client
    SUPABASE_AVAILABLE = True
except ImportError:
    SUPABASE_AVAILABLE = False

def is_cloud():
    try:
        return "supabase_url" in st.secrets
    except:
        return False

# ─────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────
st.set_page_config(
    page_title="Torre de Control · FL Servital",
    layout="wide",
    page_icon="🔧",
    initial_sidebar_state="collapsed"
)

# CSS personalizado — limpio, industrial, fácil de leer
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Header principal */
.main-header {
    background: #1a1a2e;
    color: white;
    padding: 16px 24px;
    border-radius: 12px;
    margin-bottom: 20px;
    display: flex;
    align-items: center;
    gap: 12px;
}
.main-header h1 { margin: 0; font-size: 22px; font-weight: 700; }
.main-header p  { margin: 0; font-size: 13px; opacity: 0.6; }

/* Tarjetas de operario */
.operario-card {
    background: white;
    border-radius: 12px;
    padding: 16px;
    border: 2px solid #e8e8e8;
    margin-bottom: 12px;
    transition: border-color 0.2s;
}
.operario-name {
    font-size: 13px;
    font-weight: 700;
    color: #666;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 8px;
}
.operario-task {
    font-size: 15px;
    font-weight: 600;
    color: #1a1a2e;
}
.operario-sub {
    font-size: 12px;
    color: #888;
    font-family: 'DM Mono', monospace;
}

/* Badges de estado */
.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.3px;
}
.badge-des  { background: #FFF3CD; color: #856404; }
.badge-pin  { background: #CCE5FF; color: #004085; }
.badge-arm  { background: #FFE0CC; color: #7a3a00; }
.badge-pul  { background: #E8D5FF; color: #4a0080; }
.badge-term { background: #D4EDDA; color: #155724; }
.badge-lav  { background: #D1ECF1; color: #0c5460; }
.badge-ok   { background: #D4EDDA; color: #155724; }
.badge-none { background: #f0f0f0; color: #888; }

/* Barra de progreso */
.progress-container {
    background: #f0f0f0;
    border-radius: 6px;
    height: 8px;
    overflow: hidden;
    margin: 6px 0;
}
.progress-bar {
    height: 100%;
    border-radius: 6px;
    transition: width 0.3s;
}

/* Fila del panel */
.panel-row {
    background: white;
    border-radius: 10px;
    padding: 12px 16px;
    margin-bottom: 8px;
    border: 1px solid #e8e8e8;
    display: flex;
    align-items: center;
    gap: 12px;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    background: #f5f5f5;
    padding: 4px;
    border-radius: 10px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    font-weight: 600;
    font-size: 14px;
}

/* Botón avanzar */
.stButton button {
    border-radius: 8px;
    font-weight: 600;
}

/* Métrica */
.metric-box {
    background: white;
    border-radius: 10px;
    padding: 16px;
    text-align: center;
    border: 1px solid #e8e8e8;
}
.metric-num { font-size: 32px; font-weight: 700; color: #1a1a2e; }
.metric-label { font-size: 12px; color: #888; text-transform: uppercase; letter-spacing: 0.5px; }

/* Alerta urgente */
.alerta-urgente {
    background: #fff5f5;
    border-left: 4px solid #e53e3e;
    padding: 10px 14px;
    border-radius: 6px;
    margin-bottom: 8px;
    font-size: 13px;
}

div[data-testid="stHorizontalBlock"] { gap: 12px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────
ETAPAS = ['desabolladura', 'pintura', 'armado', 'pulido', 'terminaciones', 'lavado']
ETAPA_LABEL = {
    'desabolladura': 'Desabolladura',
    'pintura':       'Pintura',
    'armado':        'Armado',
    'pulido':        'Pulido',
    'terminaciones': 'Terminaciones',
    'lavado':        'Lavado',
}
ETAPA_BADGE = {
    'desabolladura': 'des', 'pintura': 'pin', 'armado': 'arm',
    'pulido': 'pul', 'terminaciones': 'term', 'lavado': 'lav',
}
ETAPA_COLOR = {
    'desabolladura': '#856404',
    'pintura':       '#004085',
    'armado':        '#7a3a00',
    'pulido':        '#4a0080',
    'terminaciones': '#155724',
    'lavado':        '#0c5460',
}
ETAPA_PROGRESS = {
    'desabolladura': 1, 'pintura': 2, 'armado': 3,
    'pulido': 4, 'terminaciones': 5, 'lavado': 6,
}
SIGUIENTE_ETAPA = {
    'desabolladura': 'pintura',
    'pintura':       'armado',
    'armado':        'pulido',
    'pulido':        'terminaciones',
    'terminaciones': 'lavado',
    'lavado':        None,
}
ESTADO_MAP = {
    'DES': 'desabolladura', 'PIN': 'pintura', 'ARM': 'armado',
    'PUL': 'pulido', 'LAV': 'lavado', 'OK.': 'listo',
}
DES_MAP = {'RO': 'Rojas', 'GO': 'González', 'CA': 'Carvajal', 'MEC': 'Externo', 'EXT': 'Externo'}
PIN_MAP = {'AS': 'Percy', 'HU': 'Huaiquifil', 'M': 'Percy', 'ME': 'Huaiquifil', 'PE': 'Percy'}

DESABOLLADORES = ['Rojas', 'González', 'Carvajal']
PINTORES       = ['Percy', 'Huaiquifil']

DURACIONES = {
    'LEVE':    {'desabolladura': 1, 'pintura': 1, 'armado': 1, 'pulido': 1, 'terminaciones': 1, 'lavado': 1},
    'MEDIANO': {'desabolladura': 2, 'pintura': 2, 'armado': 1, 'pulido': 1, 'terminaciones': 1, 'lavado': 1},
    'GRAVE':   {'desabolladura': 5, 'pintura': 5, 'armado': 3, 'pulido': 2, 'terminaciones': 2, 'lavado': 1},
}

FERIADOS = [
    date(2026, 1, 1), date(2026, 4, 3), date(2026, 5, 1),
    date(2026, 9, 18), date(2026, 9, 19), date(2026, 12, 25)
]

hoy = date.today()

# ── Pesos de asignación por defecto (% por operario y tipo de daño)
PESOS_DEFAULT = {
    'desabolladores': {
        'LEVE':    {'Rojas': 34, 'González': 33, 'Carvajal': 33},
        'MEDIANO': {'Rojas': 34, 'González': 33, 'Carvajal': 33},
        'GRAVE':   {'Rojas': 34, 'González': 33, 'Carvajal': 33},
    },
    'pintores': {
        'LEVE':    {'Percy': 50, 'Huaiquifil': 50},
        'MEDIANO': {'Percy': 50, 'Huaiquifil': 50},
        'GRAVE':   {'Percy': 50, 'Huaiquifil': 50},
    }
}

def get_pesos():
    if 'pesos_asignacion' not in st.session_state:
        st.session_state.pesos_asignacion = PESOS_DEFAULT.copy()
    return st.session_state.pesos_asignacion

# Posiciones del taller
POSICIONES_CARROCERIA = [f"Carr{i}" for i in range(1,17)]
POSICIONES_ESP_PINTURA = [f"P{i}" for i in range(1,16)]
POSICIONES_ESP_ARMADO  = [f"Est{i}" for i in range(11,19)]
POSICIONES_ESP_PULIDO  = [f"Est{i}" for i in range(19,31)]
POSICIONES_LAVADO      = ["Lav1","Lav2","Lav3"]
POSICIONES_PRE_ENTREGA = ["PE1","PE2","PE3","PE4"]
POSICIONES_PINTURA     = [f"Pint{i}" for i in range(1,17)]
POSICIONES_HORNOS      = ["H1","H2","H3"]
TODAS_POSICIONES = (["— Sin asignar —"] +
    POSICIONES_CARROCERIA + POSICIONES_ESP_PINTURA +
    POSICIONES_ESP_ARMADO + POSICIONES_ESP_PULIDO +
    POSICIONES_LAVADO + POSICIONES_PRE_ENTREGA +
    POSICIONES_PINTURA + POSICIONES_HORNOS)

# ─────────────────────────────────────────────────────
# BASE DE DATOS (SQLite)
# ─────────────────────────────────────────────────────
import sqlite3

@st.cache_resource
def get_db():
    conn = sqlite3.connect("taller_v2.db", check_same_thread=False)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS ots (
        ot TEXT PRIMARY KEY,
        nombre TEXT, cia TEXT, ase TEXT,
        modelo TEXT, color TEXT, patente TEXT,
        dano TEXT,
        desabollador TEXT, pintor TEXT,
        etapa_actual TEXT,
        fecha_ingreso TEXT, fecha_entrega TEXT,
        comentario TEXT, repuestos TEXT, telefono TEXT,
        programacion_json TEXT,
        entregado INTEGER DEFAULT 0,
        posicion TEXT
    )''')
    try: c.execute('ALTER TABLE ots ADD COLUMN posicion TEXT')
    except: pass
    conn.commit()
    return conn

conn = get_db()

# Inicializar Supabase
def _init_supa():
    if not SUPABASE_AVAILABLE: return None
    try:
        url = st.secrets.get("supabase_url","")
        key = st.secrets.get("supabase_key","")
        if url and key:
            return create_client(url, key)
    except:
        pass
    return None

supa = _init_supa()

def save_ot(data):
    prog = json.dumps(data.get('programacion', {}))
    if supa:
        supa.table("ots").upsert({
            'ot': data['ot'], 'nombre': data.get('nombre',''),
            'cia': data.get('cia',''), 'ase': data.get('ase',''),
            'modelo': data['modelo'], 'color': data.get('color',''),
            'patente': data['patente'], 'dano': data.get('dano','MEDIANO'),
            'desabollador': data['desabollador'], 'pintor': data['pintor'],
            'etapa_actual': data['etapa_actual'],
            'fecha_ingreso': data.get('fecha_ingreso',''),
            'fecha_entrega': data.get('fecha_entrega',''),
            'comentario': data.get('comentario',''),
            'repuestos': data.get('repuestos',''),
            'telefono': data.get('telefono',''),
            'programacion_json': prog,
            'entregado': 0, 'posicion': data.get('posicion','')
        }).execute()
    else:
        c = conn.cursor()
        c.execute('''INSERT OR REPLACE INTO ots VALUES
            (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (
            data['ot'], data.get('nombre',''), data.get('cia',''), data.get('ase',''),
            data['modelo'], data.get('color',''), data['patente'],
            data.get('dano','MEDIANO'), data['desabollador'], data['pintor'],
            data['etapa_actual'], data.get('fecha_ingreso',''), data.get('fecha_entrega',''),
            data.get('comentario',''), data.get('repuestos',''), data.get('telefono',''),
            prog, 0, data.get('posicion','')
        ))
        conn.commit()

def load_ots():
    if supa:
        resp = supa.table("ots").select("*").eq("entregado", 0).order("fecha_ingreso").execute()
        rows = []
        for r in resp.data:
            try: r['programacion'] = json.loads(r.get('programacion_json') or '{}')
            except: r['programacion'] = {}
            rows.append(r)
        return rows
    else:
        c = conn.cursor()
        c.execute("SELECT * FROM ots WHERE entregado=0 ORDER BY fecha_ingreso")
        cols = ['ot','nombre','cia','ase','modelo','color','patente','dano',
                'desabollador','pintor','etapa_actual','fecha_ingreso','fecha_entrega',
                'comentario','repuestos','telefono','programacion_json','entregado','posicion']
        rows = []
        for row in c.fetchall():
            d = dict(zip(cols, row))
            try: d['programacion'] = json.loads(d['programacion_json'])
            except: d['programacion'] = {}
            rows.append(d)
        return rows

def marcar_entregado(ot):
    if supa:
        supa.table("ots").update({"entregado": 1}).eq("ot", ot).execute()
    else:
        c = conn.cursor()
        c.execute("UPDATE ots SET entregado=1 WHERE ot=?", (ot,))
        conn.commit()

def avanzar_etapa(ot, nueva_etapa):
    if supa:
        supa.table("ots").update({"etapa_actual": nueva_etapa}).eq("ot", ot).execute()
    else:
        c = conn.cursor()
        c.execute("UPDATE ots SET etapa_actual=? WHERE ot=?", (nueva_etapa, ot))
        conn.commit()

def update_ot_field(ot, field, value):
    if supa:
        supa.table("ots").update({field: value}).eq("ot", ot).execute()
    else:
        c = conn.cursor()
        c.execute(f"UPDATE ots SET {field}=? WHERE ot=?", (value, ot))
        conn.commit()

# ─────────────────────────────────────────────────────
# FUNCIONES DE PROGRAMACIÓN
# ─────────────────────────────────────────────────────
def es_laboral(d):
    return d.weekday() < 5 and d not in FERIADOS

def dias_laborales_desde(inicio, n):
    """Retorna la fecha fin después de n días laborales desde inicio"""
    d = inicio
    count = 0
    while count < n:
        if es_laboral(d):
            count += 1
        if count < n:
            d += timedelta(days=1)
    return d

def programar_ot(desabollador, pintor, dano, fecha_inicio, ots_existentes):
    """Calcula fechas de cada etapa respetando ocupación de operarios"""
    dano_key = dano.upper() if dano else 'MEDIANO'
    durs = DURACIONES.get(dano_key, DURACIONES['MEDIANO'])

    # Construir calendarios de ocupación
    cal_des = {}  # {fecha: n_ots}
    cal_pin = {}
    for ot in ots_existentes:
        prog = ot.get('programacion', {})
        des = ot.get('desabollador', '')
        pin = ot.get('pintor', '')
        for etapa, info in prog.items():
            if not info: continue
            try:
                fi = datetime.strptime(info['inicio'], '%Y-%m-%d').date()
                ff = datetime.strptime(info['fin'], '%Y-%m-%d').date()
            except: continue
            d = fi
            while d <= ff:
                if es_laboral(d):
                    if etapa in ('desabolladura', 'armado') and des == desabollador:
                        cal_des[d] = cal_des.get(d, 0) + 1
                    if etapa in ('pintura', 'pulido') and pin == pintor:
                        cap = 2 if 'Huaiquifil' in pin else 1
                        cal_pin[d] = cal_pin.get(d, 0) + 1
                d += timedelta(days=1)

    def buscar_inicio_libre(cal, desde, dur, capacidad=1):
        d = max(desde, hoy)
        while True:
            ok = True
            temp = d
            dias = 0
            while dias < dur:
                if es_laboral(temp):
                    if cal.get(temp, 0) >= capacidad:
                        ok = False
                        break
                    dias += 1
                temp += timedelta(days=1)
            if ok:
                return d
            d += timedelta(days=1)

    cap_pin = 2 if 'Huaiquifil' in pintor else 1
    prog = {}
    auto_desde = fecha_inicio

    for etapa in ETAPAS:
        dur = durs[etapa]
        if etapa in ('terminaciones', 'lavado'):
            inicio = auto_desde
            fin = dias_laborales_desde(inicio, dur)
            prog[etapa] = {'inicio': inicio.strftime('%Y-%m-%d'), 'fin': fin.strftime('%Y-%m-%d')}
        elif etapa in ('desabolladura', 'armado'):
            inicio = buscar_inicio_libre(cal_des, auto_desde, dur, 1)
            fin = dias_laborales_desde(inicio, dur)
            # Bloquear días usados
            d = inicio
            cnt = 0
            while cnt < dur:
                if es_laboral(d):
                    cal_des[d] = cal_des.get(d, 0) + 1
                    cnt += 1
                d += timedelta(days=1)
            prog[etapa] = {'inicio': inicio.strftime('%Y-%m-%d'), 'fin': fin.strftime('%Y-%m-%d')}
        else:  # pintura, pulido
            inicio = buscar_inicio_libre(cal_pin, auto_desde, dur, cap_pin)
            fin = dias_laborales_desde(inicio, dur)
            d = inicio
            cnt = 0
            while cnt < dur:
                if es_laboral(d):
                    cal_pin[d] = cal_pin.get(d, 0) + 1
                    cnt += 1
                d += timedelta(days=1)
            prog[etapa] = {'inicio': inicio.strftime('%Y-%m-%d'), 'fin': fin.strftime('%Y-%m-%d')}

        fin_date = datetime.strptime(prog[etapa]['fin'], '%Y-%m-%d').date()
        auto_desde = fin_date + timedelta(days=1)

    return prog

def mejor_desabollador(ots, dano='MEDIANO'):
    """Elige el desabollador según carga y pesos configurados"""
    pesos = get_pesos()['desabolladores'].get(dano.upper(), PESOS_DEFAULT['desabolladores']['MEDIANO'])
    carga = {d: 0.0 for d in DESABOLLADORES}
    fin_14 = hoy + timedelta(days=14)
    for ot in ots:
        prog = ot.get('programacion', {})
        des = ot.get('desabollador', '')
        if des not in DESABOLLADORES: continue
        for etapa in ['desabolladura', 'armado']:
            info = prog.get(etapa)
            if not info: continue
            try:
                fi = datetime.strptime(info['inicio'], '%Y-%m-%d').date()
                ff = datetime.strptime(info['fin'], '%Y-%m-%d').date()
                if fi <= fin_14:
                    carga[des] += (min(ff, fin_14) - fi).days + 1
            except: pass
    # Ajustar carga según pesos: menor peso = más prioridad
    # Score = carga / peso (a mayor peso, se tolera más carga)
    score = {}
    for d in DESABOLLADORES:
        peso = pesos.get(d, 33)
        score[d] = carga[d] / max(peso, 1)
    return min(score, key=score.get)

def mejor_pintor(ots, dano='MEDIANO'):
    """Elige el pintor según carga, capacidad y pesos configurados"""
    pesos = get_pesos()['pintores'].get(dano.upper(), PESOS_DEFAULT['pintores']['MEDIANO'])
    carga = {p: 0.0 for p in PINTORES}
    fin_14 = hoy + timedelta(days=14)
    for ot in ots:
        prog = ot.get('programacion', {})
        pin = ot.get('pintor', '')
        if pin not in PINTORES: continue
        cap = 2 if pin == 'Huaiquifil' else 1
        for etapa in ['pintura', 'pulido']:
            info = prog.get(etapa)
            if not info: continue
            try:
                fi = datetime.strptime(info['inicio'], '%Y-%m-%d').date()
                ff = datetime.strptime(info['fin'], '%Y-%m-%d').date()
                if fi <= fin_14:
                    carga[pin] += ((min(ff, fin_14) - fi).days + 1) / cap
            except: pass
    score = {}
    for p in PINTORES:
        peso = pesos.get(p, 50)
        score[p] = carga[p] / max(peso, 1)
    return min(score, key=score.get)

# ─────────────────────────────────────────────────────
# HELPERS UI
# ─────────────────────────────────────────────────────
def badge_etapa(etapa):
    if not etapa or etapa == 'listo':
        return "<span class='badge badge-ok'>✅ Listo</span>"
    b = ETAPA_BADGE.get(etapa, 'none')
    label = ETAPA_LABEL.get(etapa, etapa)
    return f"<span class='badge badge-{b}'>{label}</span>"

def progress_pct(etapa):
    if etapa == 'listo': return 100
    return round(ETAPA_PROGRESS.get(etapa, 0) / 6 * 100)

def progress_color(etapa):
    colors = {
        'desabolladura': '#F6AD55', 'pintura': '#63B3ED',
        'armado': '#FC8181', 'pulido': '#B794F4',
        'terminaciones': '#68D391', 'lavado': '#4FD1C5',
        'listo': '#48BB78'
    }
    return colors.get(etapa, '#CBD5E0')

def dias_restantes(fecha_str):
    if not fecha_str: return None
    try:
        f = datetime.strptime(fecha_str, '%Y-%m-%d').date()
        return (f - hoy).days
    except: return None

def format_fecha(fecha_str):
    if not fecha_str: return "—"
    try:
        return datetime.strptime(fecha_str, '%Y-%m-%d').strftime('%d/%m/%y')
    except: return fecha_str

# ─────────────────────────────────────────────────────
# IMPORTAR DESDE PLANILLA (primera vez)
# ─────────────────────────────────────────────────────
def importar_planilla(df_raw):
    importados = 0
    ots_actuales = {ot['ot'] for ot in load_ots()}

    for _, row in df_raw.iterrows():
        try:
            ot_num = str(row.get('O T', '')).strip().replace('.0','')
            if not ot_num or ot_num == 'nan': continue
            ot_id = f"OT-{ot_num}"
            if ot_id in ots_actuales: continue

            estado_raw = str(row.get('EST', '')).strip().upper()
            etapa = ESTADO_MAP.get(estado_raw, 'desabolladura')
            if etapa == 'listo': continue  # ya entregados no se importan

            des_raw = str(row.get('DES', '')).strip().upper()
            pin_raw = str(row.get('PIN', '')).strip().upper()
            des = DES_MAP.get(des_raw, 'Rojas')
            pin = PIN_MAP.get(pin_raw, 'Percy')

            # LLEGA = fecha ingreso al taller (base para programar etapas)
            try:
                fi_raw = pd.to_datetime(row.get('LLEGA'))
                if pd.isna(fi_raw): raise ValueError
                fi = fi_raw.date()
                fecha_tentativa = False
            except:
                fi = hoy
                fecha_tentativa = True

            # TALLER = fecha objetivo interna (cuándo debe estar listo)
            try:
                ft_raw = pd.to_datetime(row.get('TALLER'))
                ft_str = ft_raw.strftime('%Y-%m-%d') if not pd.isna(ft_raw) else ''
            except:
                ft_str = ''

            dano_raw = str(row.get('DAÑO', 'MEDIANO')).strip().upper()
            if dano_raw not in ('LEVE','MEDIANO','GRAVE'):
                dano_raw = 'MEDIANO'

            ots_actual = load_ots()
            prog = programar_ot(des, pin, dano_raw, fi, ots_actual)

            data = {
                'ot': ot_id,
                'nombre': str(row.get('NOMBRE', '') or '').strip(),
                'cia': str(row.get('CIA', '') or '').strip(),
                'ase': str(row.get('ASE', '') or '').strip(),
                'modelo': str(row.get('MODELO', '') or '').strip(),
                'color': str(row.get('COLOR', '') or '').strip(),
                'patente': str(row.get('PATENTE', '') or '').strip(),
                'dano': dano_raw,
                'desabollador': des,
                'pintor': pin,
                'etapa_actual': etapa,
                'fecha_ingreso': fi.strftime('%Y-%m-%d'),
                'fecha_entrega': ft_str,
                'comentario': ('⚠️ Fecha tentativa | ' if fecha_tentativa else '') + str(row.get('COMENTARIO', '') or '').strip(),
                'repuestos': str(row.get('REPUESTOS', '') or '').strip(),
                'telefono': str(row.get('TELEFONO', '') or '').strip(),
                'programacion': prog,
            }
            save_ot(data)
            importados += 1
        except Exception as e:
            continue
    return importados

# ─────────────────────────────────────────────────────
# EXPORTAR A EXCEL
# ─────────────────────────────────────────────────────
def exportar_excel(ots_data):
    ETAPAS_SEQ = ['desabolladura','pintura','armado','pulido','terminaciones','lavado']
    ETAPA_IDX  = {e: i for i, e in enumerate(ETAPAS_SEQ)}
    DES_INV = {'Rojas':'RO','González':'GO','Carvajal':'CA','Externo':'EXT'}
    PIN_INV = {'Percy':'AS','Huaiquifil':'HU'}
    ETAPA_ABREV = {'desabolladura':'DES','pintura':'PIN','armado':'ARM',
                   'pulido':'PUL','terminaciones':'TERM','lavado':'LAV','listo':'OK.'}

    ROJO       = PatternFill("solid", fgColor="C0131A")
    GRIS_OSC   = PatternFill("solid", fgColor="404040")
    GRIS_CLAR  = PatternFill("solid", fgColor="F2F2F2")
    AMARILLO   = PatternFill("solid", fgColor="FFD966")
    VERDE_F    = PatternFill("solid", fgColor="C6EFCE")
    AZUL_F     = PatternFill("solid", fgColor="DDEEFF")
    NARANJO    = PatternFill("solid", fgColor="FCE4D6")
    MORADO     = PatternFill("solid", fgColor="E8D5FF")
    BLANCO_F   = PatternFill("solid", fgColor="FFFFFF")

    borde = Border(left=Side(style='thin'), right=Side(style='thin'),
                   top=Side(style='thin'), bottom=Side(style='thin'))
    centro = Alignment(horizontal='center', vertical='center', wrap_text=True)
    izq    = Alignment(horizontal='left', vertical='center', wrap_text=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "EN TRABAJO"

    # Título
    ws.merge_cells('A1:Z1')
    ws['A1'] = "PLANILLA D&P — FL SERVITAL"
    ws['A1'].font = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    ws['A1'].fill = ROJO
    ws['A1'].alignment = centro
    ws.row_dimensions[1].height = 22

    ws.merge_cells('A2:Z2')
    ws['A2'] = f"Exportado: {datetime.now().strftime('%d/%m/%Y %H:%M')}  ·  {len(ots_data)} vehículos"
    ws['A2'].font = Font(name="Arial", color="FFFFFF", size=8)
    ws['A2'].fill = PatternFill("solid", fgColor="404040")
    ws['A2'].alignment = centro
    ws.row_dimensions[2].height = 14

    headers = [
        ('N°',8),('NOMBRE',22),('CIA',8),('ASE',5),('O T',10),
        ('MODELO',12),('COLOR',8),('PATENTE',10),('DAÑO',9),
        ('DES',6),('PIN',6),
        ('DES✓',6),('PIN✓',6),('ARM✓',6),('PUL✓',6),('LAV✓',6),('OK.',6),
        ('EST',9),('POSICIÓN',10),
        ('COMENTARIO',28),('REPUESTOS',22),
        ('TALLER',10),('CLIENTE',10),('TELÉFONO',13),
    ]
    FILL_H = [PatternFill("solid", fgColor="404040")]*11 +              [AMARILLO,AZUL_F,NARANJO,MORADO,VERDE_F,VERDE_F] +              [PatternFill("solid", fgColor="C0131A")] +              [PatternFill("solid", fgColor="404040")]*6

    for ci, ((h, w), fh) in enumerate(zip(headers, FILL_H), 1):
        cell = ws.cell(row=3, column=ci, value=h)
        cell.fill = fh
        tc = "FFFFFF" if fh.fgColor.rgb in ("FF404040","FFC0131A") else "000000"
        cell.font = Font(name="Arial", bold=True, color=tc, size=9)
        cell.alignment = centro
        cell.border = borde
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[3].height = 18

    COLOR_EST = {
        'desabolladura': AMARILLO, 'pintura': AZUL_F, 'armado': NARANJO,
        'pulido': MORADO, 'terminaciones': VERDE_F, 'lavado': VERDE_F, 'listo': VERDE_F
    }

    def fdate(s):
        if not s: return ''
        try: return datetime.strptime(s, '%Y-%m-%d').strftime('%d/%m/%Y')
        except: return s

    for rn, ot in enumerate(ots_data, 4):
        etapa = ot.get('etapa_actual','')
        ei = ETAPA_IDX.get(etapa, -1)
        fill_r = COLOR_EST.get(etapa, BLANCO_F)

        def marca(e):
            if etapa == 'listo': return 'X'
            return 'X' if ETAPA_IDX.get(e,-1) < ei else ''

        vals = [
            rn-3,
            ot.get('nombre',''), ot.get('cia',''), ot.get('ase',''),
            ot.get('ot','').replace('OT-',''),
            ot.get('modelo',''), ot.get('color',''), ot.get('patente',''),
            ot.get('dano',''),
            DES_INV.get(ot.get('desabollador',''), ot.get('desabollador','')),
            PIN_INV.get(ot.get('pintor',''), ot.get('pintor','')),
            marca('desabolladura'), marca('pintura'), marca('armado'),
            marca('pulido'), marca('lavado'),
            'X' if etapa == 'listo' else '',
            ETAPA_ABREV.get(etapa, etapa.upper()[:3]),
            ot.get('posicion',''),
            ot.get('comentario',''), ot.get('repuestos',''),
            fdate(ot.get('fecha_ingreso','')),
            fdate(ot.get('fecha_entrega','')),
            ot.get('telefono',''),
        ]

        for ci, val in enumerate(vals, 1):
            cell = ws.cell(row=rn, column=ci, value=val)
            cell.border = borde
            if ci == 18:  # EST
                cell.fill = fill_r
                cell.font = Font(name="Arial", bold=True, color="C0131A", size=9)
                cell.alignment = centro
            elif ci in range(12, 18):  # marcas X
                cell.fill = VERDE_F if val == 'X' else BLANCO_F
                cell.font = Font(name="Arial", bold=True,
                                 color="155724" if val == 'X' else "AAAAAA", size=9)
                cell.alignment = centro
            else:
                cell.fill = GRIS_CLAR if rn % 2 == 0 else BLANCO_F
                cell.font = Font(name="Arial", color="000000", size=9)
                cell.alignment = izq if ci in [2,20,21] else centro
        ws.row_dimensions[rn].height = 16

    ws.freeze_panes = 'A4'
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}3"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────
# Debug: mostrar modo de BD
if supa:
    st.sidebar.success("☁️ BD: Supabase")
else:
    st.sidebar.info("💻 BD: Local (SQLite)")

ots = load_ots()
n_urgentes = sum(1 for o in ots if dias_restantes(o.get('fecha_entrega','')) is not None
                 and dias_restantes(o.get('fecha_entrega','')) <= 3
                 and o['etapa_actual'] != 'listo')

st.markdown(f"""
<div class='main-header'>
    <div style='font-size:28px'>🔧</div>
    <div>
        <h1>Torre de Control · FL Servital</h1>
        <p>{hoy.strftime('%A %d de %B, %Y').capitalize()} &nbsp;·&nbsp; {len(ots)} vehículos activos
        {'&nbsp;·&nbsp; <span style="color:#ff6b6b">⚠️ ' + str(n_urgentes) + ' urgentes</span>' if n_urgentes else ''}</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────
tab_hoy, tab_panel, tab_crono, tab_nueva, tab_config = st.tabs([
    "📅  Hoy", "📋  Panel", "📊  Cronograma", "➕  Nueva OT", "⚙️  Config"
])

# ══════════════════════════════════════════════════════
# TAB 1 — HOY
# ══════════════════════════════════════════════════════
with tab_hoy:
    st.markdown(f"### Trabajos del día — {hoy.strftime('%d/%m/%Y')}")

    # Encontrar qué hace cada operario HOY
    def trabajo_hoy(nombre_operario, tipo):
        trabajos = []
        for ot in ots:
            prog = ot.get('programacion', {})
            operario = ot['desabollador'] if tipo == 'des' else ot['pintor']
            if operario != nombre_operario: continue
            for etapa, info in prog.items():
                if not info: continue
                try:
                    fi = datetime.strptime(info['inicio'], '%Y-%m-%d').date()
                    ff = datetime.strptime(info['fin'], '%Y-%m-%d').date()
                    if fi <= hoy <= ff:
                        trabajos.append({'ot': ot['ot'], 'patente': ot['patente'],
                                         'modelo': ot['modelo'], 'etapa': etapa,
                                         'etapa_actual': ot['etapa_actual']})
                except: pass
        return trabajos

    # Alertas urgentes
    urgentes = [o for o in ots if dias_restantes(o.get('fecha_entrega','')) is not None
                and dias_restantes(o.get('fecha_entrega','')) <= 3
                and o['etapa_actual'] != 'listo']
    if urgentes:
        st.markdown("#### 🚨 Urgentes")
        for u in urgentes:
            dr = dias_restantes(u['fecha_entrega'])
            color = "#c53030" if dr < 0 else "#dd6b20"
            msg = f"{'ATRASADO ' + str(abs(dr)) + ' días' if dr < 0 else 'Entrega en ' + str(dr) + ' días'}"
            st.markdown(f"""
            <div class='alerta-urgente'>
                <b>{u['ot']}</b> · {u['patente']} · {u['modelo']}
                &nbsp;{badge_etapa(u['etapa_actual'])}
                &nbsp;<span style='color:{color};font-weight:700'>{msg}</span>
                &nbsp;· Des: <b>{u['desabollador']}</b> · Pin: <b>{u['pintor']}</b>
            </div>
            """, unsafe_allow_html=True)
        st.markdown("---")

    # Tarjetas por operario
    st.markdown("#### 👷 Desabolladores")
    cols_des = st.columns(3)
    for i, des in enumerate(DESABOLLADORES):
        trabajos = trabajo_hoy(des, 'des')
        carga = len(trabajos)
        color_borde = "#48BB78" if carga == 0 else "#F6AD55" if carga <= 2 else "#FC8181"
        with cols_des[i]:
            if trabajos:
                tareas_html = "".join([
                    f"<div style='margin-bottom:6px'>"
                    f"<span style='font-family:DM Mono,monospace;font-size:12px;color:#666'>{t['ot']} · {t['patente']}</span><br>"
                    f"<b style='font-size:14px'>{t['modelo']}</b> &nbsp; {badge_etapa(t['etapa'])}"
                    f"</div>"
                    for t in trabajos
                ])
            else:
                tareas_html = "<span style='color:#aaa;font-size:13px'>Sin trabajos hoy</span>"

            st.markdown(f"""
            <div class='operario-card' style='border-color:{color_borde}'>
                <div class='operario-name'>{des}</div>
                {tareas_html}
            </div>
            """, unsafe_allow_html=True)

    st.markdown("#### 🎨 Pintores")
    cols_pin = st.columns(2)
    for i, pin in enumerate(PINTORES):
        trabajos = trabajo_hoy(pin, 'pin')
        carga = len(trabajos)
        cap = 2 if pin == 'Huaiquifil' else 1
        color_borde = "#48BB78" if carga == 0 else "#F6AD55" if carga < cap else "#FC8181"
        with cols_pin[i]:
            if trabajos:
                tareas_html = "".join([
                    f"<div style='margin-bottom:6px'>"
                    f"<span style='font-family:DM Mono,monospace;font-size:12px;color:#666'>{t['ot']} · {t['patente']}</span><br>"
                    f"<b style='font-size:14px'>{t['modelo']}</b> &nbsp; {badge_etapa(t['etapa'])}"
                    f"</div>"
                    for t in trabajos
                ])
            else:
                tareas_html = "<span style='color:#aaa;font-size:13px'>Sin trabajos hoy</span>"

            st.markdown(f"""
            <div class='operario-card' style='border-color:{color_borde}'>
                <div class='operario-name'>{pin} <span style='font-size:11px;color:#999'>(cap. {cap})</span></div>
                {tareas_html}
            </div>
            """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
# TAB 2 — PANEL DE PROGRAMACIÓN
# ══════════════════════════════════════════════════════
with tab_panel:
    # Filtros rápidos
    f1, f2, f3, f4 = st.columns([2, 2, 2, 2])
    filtro_texto  = f1.text_input("🔍 Buscar patente / modelo / OT", key="p_buscar")
    filtro_etapa  = f2.selectbox("Etapa", ["Todas"] + [ETAPA_LABEL[e] for e in ETAPAS], key="p_etapa")
    filtro_des    = f3.selectbox("Desabollador", ["Todos"] + DESABOLLADORES, key="p_des")
    filtro_pin    = f4.selectbox("Pintor", ["Todos"] + PINTORES, key="p_pin")

    ots_filtradas = ots
    if filtro_texto:
        txt = filtro_texto.lower()
        ots_filtradas = [o for o in ots_filtradas if
            txt in o['patente'].lower() or txt in o['modelo'].lower() or txt in o['ot'].lower()]
    if filtro_etapa != "Todas":
        etapa_key = [k for k, v in ETAPA_LABEL.items() if v == filtro_etapa]
        if etapa_key:
            ots_filtradas = [o for o in ots_filtradas if o['etapa_actual'] == etapa_key[0]]
    if filtro_des != "Todos":
        ots_filtradas = [o for o in ots_filtradas if o['desabollador'] == filtro_des]
    if filtro_pin != "Todos":
        ots_filtradas = [o for o in ots_filtradas if o['pintor'] == filtro_pin]

    col_cap, col_dl = st.columns([3, 1])
    col_cap.caption(f"Mostrando {len(ots_filtradas)} de {len(ots)} vehículos")
    excel_buf = exportar_excel(ots)
    col_dl.download_button("📥 Descargar Excel", excel_buf,
        file_name=f"Planilla_Taller_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)
    st.markdown("---")

    # Panel OT por OT
    if 'editar_ot' not in st.session_state:
        st.session_state.editar_ot = None

    for ot in ots_filtradas:
        # Fecha taller (ingreso) es la base para atrasos y urgencias
        fi_fmt = format_fecha(ot.get('fecha_ingreso',''))
        fe_fmt = format_fecha(ot.get('fecha_entrega',''))
        dr = dias_restantes(ot.get('fecha_ingreso',''))
        pct = progress_pct(ot['etapa_actual'])
        pcolor = progress_color(ot['etapa_actual'])
        es_listo = ot['etapa_actual'] == 'listo'
        tentativa = ot.get('comentario','').startswith('⚠️ Fecha tentativa')

        # Color fondo según urgencia
        if not es_listo and dr is not None and dr < 0:
            row_bg = "#fff5f5"
        elif not es_listo and dr is not None and dr <= 3:
            row_bg = "#fffbf0"
        else:
            row_bg = "white"

        # Fecha label
        if tentativa:
            fecha_label = "<span style='color:#e07b00;font-weight:600'>⚠️ Fecha tentativa</span>"
        elif fi_fmt and fi_fmt != '—':
            if not es_listo and dr is not None and dr < 0:
                fecha_label = f"<span style='color:#c53030;font-weight:700'>🚨 Taller: {fi_fmt} ({abs(dr)}d atraso)</span>"
            elif not es_listo and dr is not None and dr <= 3:
                fecha_label = f"<span style='color:#dd6b20;font-weight:700'>⚠️ Taller: {fi_fmt} ({dr}d)</span>"
            else:
                fecha_label = f"<span style='color:#444'>🔧 Taller: {fi_fmt}</span>"
            if fe_fmt and fe_fmt != '—':
                fecha_label += f" &nbsp;<span style='color:#999;font-size:11px'>· Cliente: {fe_fmt}</span>"
        else:
            fecha_label = "<span style='color:#e07b00;font-weight:600'>⚠️ Fecha tentativa</span>"

        posicion_html = f" · 📍 <b>{ot.get('posicion','')}</b>" if ot.get('posicion') else ''

        col_info, col_prog, col_acc = st.columns([3, 4, 2])

        with col_info:
            st.markdown(f"""
            <div style='background:{row_bg};padding:12px;border-radius:10px;border:1px solid #e8e8e8;height:100%'>
                <div style='font-size:11px;color:#888;font-family:DM Mono,monospace'>{ot['ot']}</div>
                <div style='font-size:16px;font-weight:700;color:#1a1a2e'>{ot['patente']} · {ot['modelo']}</div>
                <div style='font-size:12px;color:#666;margin-top:2px'>{ot.get('nombre') or '—'} · {ot.get('cia') or '—'}</div>
                <div style='margin-top:6px;font-size:12px'>{fecha_label}</div>
                <div style='font-size:12px;color:#888;margin-top:4px'>Des: <b>{ot['desabollador']}</b> · Pin: <b>{ot['pintor']}</b>{posicion_html}</div>
            </div>
            """, unsafe_allow_html=True)

        with col_prog:
            st.markdown(f"""
            <div style='background:{row_bg};padding:12px;border-radius:10px;border:1px solid #e8e8e8;height:100%'>
                <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:6px'>
                    <span style='font-size:12px;color:#666'>Progreso</span>
                    {badge_etapa(ot['etapa_actual'])}
                </div>
                <div class='progress-container'>
                    <div class='progress-bar' style='width:{pct}%;background:{pcolor}'></div>
                </div>
                <div style='display:flex;justify-content:space-between;margin-top:8px'>
                    {''.join([
                        f"<span style='font-size:10px;color:{"#1a1a2e" if ETAPA_PROGRESS.get(ot["etapa_actual"],0) > ETAPA_PROGRESS.get(e,0) else ("#2B6CB0" if e == ot["etapa_actual"] else "#ccc")};font-weight:{"700" if e == ot["etapa_actual"] else "400"}'>{ETAPA_LABEL[e][:3]}</span>"
                        for e in ETAPAS
                    ])}
                </div>
                {'<div style="font-size:11px;color:#888;margin-top:6px">💬 ' + ot["comentario"][:60] + ('...' if len(ot.get("comentario","")) > 60 else '') + '</div>' if ot.get("comentario") else ""}
            </div>
            """, unsafe_allow_html=True)

        with col_acc:
            sig = SIGUIENTE_ETAPA.get(ot['etapa_actual'])
            if sig:
                if st.button(f"→ {ETAPA_LABEL[sig]}", key=f"av_{ot['ot']}", use_container_width=True, type="primary"):
                    avanzar_etapa(ot['ot'], sig)
                    st.rerun()
            else:
                if st.button("🚗 Entregar", key=f"en_{ot['ot']}", use_container_width=True):
                    marcar_entregado(ot['ot'])
                    st.rerun()

            if st.button("✏️ Editar", key=f"ed_{ot['ot']}", use_container_width=True):
                st.session_state.editar_ot = ot['ot'] if st.session_state.editar_ot != ot['ot'] else None
                st.rerun()

        # Panel de edición inline
        if st.session_state.editar_ot == ot['ot']:
            with st.container():
                st.markdown(f"**✏️ Editando {ot['ot']} — {ot['patente']}**")
                e1, e2, e3 = st.columns(3)
                etapa_opts = ETAPAS
                etapa_idx = etapa_opts.index(ot['etapa_actual']) if ot['etapa_actual'] in etapa_opts else 0
                nueva_etapa = e1.selectbox("Etapa actual", etapa_opts,
                    format_func=lambda x: ETAPA_LABEL[x],
                    index=etapa_idx, key=f"et_{ot['ot']}")
                nuevo_des = e2.selectbox("Desabollador", DESABOLLADORES,
                    index=DESABOLLADORES.index(ot['desabollador']) if ot['desabollador'] in DESABOLLADORES else 0,
                    key=f"des_{ot['ot']}")
                nuevo_pin = e3.selectbox("Pintor", PINTORES,
                    index=PINTORES.index(ot['pintor']) if ot['pintor'] in PINTORES else 0,
                    key=f"pin_{ot['ot']}")
                nuevo_com = st.text_area("Comentario", value=ot.get('comentario',''), height=60, key=f"com_{ot['ot']}")
                nuevo_rep = st.text_input("Repuestos", value=ot.get('repuestos',''), key=f"rep_{ot['ot']}")
                pos_actual = ot.get('posicion','') or '— Sin asignar —'
                pos_idx = TODAS_POSICIONES.index(pos_actual) if pos_actual in TODAS_POSICIONES else 0
                nueva_pos = st.selectbox("📍 Posición en el taller", TODAS_POSICIONES,
                    index=pos_idx, key=f"pos_{ot['ot']}")

                s1, s2, _ = st.columns([1, 1, 4])
                if s1.button("💾 Guardar", key=f"save_{ot['ot']}", type="primary"):
                    pos_guardar = nueva_pos if nueva_pos != '— Sin asignar —' else ''
                    if supa:
                        supa.table("ots").update({
                            'etapa_actual': nueva_etapa, 'desabollador': nuevo_des,
                            'pintor': nuevo_pin, 'comentario': nuevo_com,
                            'repuestos': nuevo_rep, 'posicion': pos_guardar
                        }).eq("ot", ot['ot']).execute()
                    else:
                        c = conn.cursor()
                        c.execute("""UPDATE ots SET etapa_actual=?, desabollador=?, pintor=?,
                                    comentario=?, repuestos=?, posicion=? WHERE ot=?""",
                                 (nueva_etapa, nuevo_des, nuevo_pin, nuevo_com, nuevo_rep, pos_guardar, ot['ot']))
                        conn.commit()
                    st.session_state.editar_ot = None
                    st.success("✅ Guardado")
                    st.rerun()
                if s2.button("✖ Cerrar", key=f"cls_{ot['ot']}"):
                    st.session_state.editar_ot = None
                    st.rerun()

        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
# TAB 3 — CRONOGRAMA
# ══════════════════════════════════════════════════════
with tab_crono:
    cr1, cr2 = st.columns([2, 3])
    vista = cr1.radio("Vista", ["Por semana", "Por desabollador", "Por pintor"], horizontal=True)
    if 'semana_offset' not in st.session_state:
        st.session_state.semana_offset = 0

    if vista == "Por semana":
        n1, n2, n3, n4 = st.columns([1, 1, 2, 2])
        if n1.button("◀ Anterior"): st.session_state.semana_offset -= 1; st.rerun()
        if n2.button("Siguiente ▶"): st.session_state.semana_offset += 1; st.rerun()
        filtro_op_crono = n3.multiselect("Filtrar operarios:",
            DESABOLLADORES + PINTORES,
            default=DESABOLLADORES + PINTORES,
            key="crono_op")

        lunes = hoy - timedelta(days=hoy.weekday()) + timedelta(weeks=st.session_state.semana_offset)
        dias = [lunes + timedelta(days=i) for i in range(5)]

        st.markdown(f"#### Semana del {lunes.strftime('%d/%m')} al {dias[-1].strftime('%d/%m/%Y')}")

        # Construir tabla
        todos_operarios = DESABOLLADORES + PINTORES
        tabla_data = {op: {d: [] for d in dias} for op in todos_operarios}

        for ot in ots:
            prog = ot.get('programacion', {})
            for etapa, info in prog.items():
                if not info: continue
                try:
                    fi = datetime.strptime(info['inicio'], '%Y-%m-%d').date()
                    ff = datetime.strptime(info['fin'], '%Y-%m-%d').date()
                except: continue
                if etapa in ('desabolladura', 'armado'):
                    op = ot['desabollador']
                elif etapa in ('pintura', 'pulido'):
                    op = ot['pintor']
                else:
                    continue
                if op not in tabla_data: continue
                for dia in dias:
                    if fi <= dia <= ff and es_laboral(dia):
                        tabla_data[op][dia].append(f"{ot['patente']} ({ETAPA_LABEL[etapa][:3]})")

        # Renderizar tabla
        cols_dias = st.columns([2] + [1]*5)
        cols_dias[0].markdown("**Operario**")
        for i, d in enumerate(dias):
            es_hoy = d == hoy
            label = f"**{'🔵 ' if es_hoy else ''}{d.strftime('%a %d/%m')}**"
            cols_dias[i+1].markdown(label)

        st.markdown("<hr style='margin:4px 0'>", unsafe_allow_html=True)

        todos_operarios_filtrados = [op for op in todos_operarios if op in filtro_op_crono]
        for op in todos_operarios_filtrados:
            cols = st.columns([2] + [1]*5)
            cols[0].markdown(f"**{op}**")
            for i, d in enumerate(dias):
                tareas = tabla_data[op][d]
                if tareas:
                    for t in tareas:
                        cols[i+1].markdown(f"<span style='font-size:12px;background:#e8f4fd;padding:2px 6px;border-radius:4px;display:block;margin-bottom:2px'>{t}</span>", unsafe_allow_html=True)
                else:
                    cols[i+1].markdown("<span style='color:#ccc;font-size:12px'>—</span>", unsafe_allow_html=True)

    elif vista == "Por desabollador":
        des_sel = st.selectbox("Seleccionar desabollador", DESABOLLADORES)
        ots_des = [o for o in ots if o['desabollador'] == des_sel]
        st.markdown(f"#### {des_sel} — {len(ots_des)} vehículos asignados")

        for ot in ots_des:
            prog = ot.get('programacion', {})
            pct = progress_pct(ot['etapa_actual'])
            pcolor = progress_color(ot['etapa_actual'])
            dr = dias_restantes(ot['fecha_entrega'])

            col_a, col_b = st.columns([2, 3])
            with col_a:
                st.markdown(f"""
                <div style='padding:10px;background:white;border-radius:8px;border:1px solid #e8e8e8'>
                    <div style='font-size:11px;color:#888'>{ot['ot']}</div>
                    <b>{ot['patente']} · {ot['modelo']}</b>
                    <div style='font-size:12px;color:#666'>Entrega: {format_fecha(ot['fecha_entrega'])}
                    {"· <span style='color:#c53030'>ATRASADO</span>" if dr is not None and dr < 0 else ""}
                    </div>
                </div>
                """, unsafe_allow_html=True)
            with col_b:
                # Mini cronograma de etapas
                etapas_html = ""
                for e in ['desabolladura', 'armado']:
                    info = prog.get(e)
                    if info:
                        fi = format_fecha(info['inicio'])
                        ff = format_fecha(info['fin'])
                        activo = ot['etapa_actual'] == e
                        bg = "#2B6CB0" if activo else "#e8f4fd"
                        tc = "white" if activo else "#2B6CB0"
                        etapas_html += f"<span style='background:{bg};color:{tc};padding:3px 8px;border-radius:4px;font-size:11px;margin-right:4px'>{ETAPA_LABEL[e]}: {fi}→{ff}</span>"
                st.markdown(f"<div style='padding:10px'>{etapas_html}</div>", unsafe_allow_html=True)

    elif vista == "Por pintor":
        pin_sel = st.selectbox("Seleccionar pintor", PINTORES)
        ots_pin = [o for o in ots if o['pintor'] == pin_sel]
        st.markdown(f"#### {pin_sel} — {len(ots_pin)} vehículos asignados")

        for ot in ots_pin:
            prog = ot.get('programacion', {})
            dr = dias_restantes(ot['fecha_entrega'])
            col_a, col_b = st.columns([2, 3])
            with col_a:
                atras = "· <span style='color:#c53030'>ATRASADO</span>" if dr is not None and dr < 0 else ""
                st.markdown(f"<div style='padding:10px;background:white;border-radius:8px;border:1px solid #e8e8e8'>"
                    f"<div style='font-size:11px;color:#888'>{ot['ot']}</div>"
                    f"<b>{ot['patente']} · {ot['modelo']}</b>"
                    f"<div style='font-size:12px;color:#666'>Entrega: {format_fecha(ot['fecha_entrega'])} {atras}</div>"
                    f"</div>", unsafe_allow_html=True)
            with col_b:
                etapas_html = ""
                for e in ['pintura', 'pulido']:
                    info = prog.get(e)
                    if info:
                        fi = format_fecha(info['inicio'])
                        ff = format_fecha(info['fin'])
                        activo = ot['etapa_actual'] == e
                        bg = "#2B6CB0" if activo else "#e8f4fd"
                        tc = "white" if activo else "#2B6CB0"
                        etapas_html += f"<span style='background:{bg};color:{tc};padding:3px 8px;border-radius:4px;font-size:11px;margin-right:4px'>{ETAPA_LABEL[e]}: {fi}→{ff}</span>"
                st.markdown(f"<div style='padding:10px'>{etapas_html}</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
# TAB 4 — NUEVA OT
# ══════════════════════════════════════════════════════
with tab_nueva:
    st.markdown("### ➕ Ingresar nuevo vehículo")
    st.markdown("El sistema asignará automáticamente el equipo más libre.")

    with st.form("form_nueva_ot"):
        st.markdown("**Datos del vehículo**")
        n1, n2, n3, n4 = st.columns(4)
        f_ot      = n1.text_input("N° OT *", placeholder="219999")
        f_patente = n2.text_input("Patente *", placeholder="XXXX99")
        f_modelo  = n3.text_input("Modelo *", placeholder="Tiggo 8")
        f_color   = n4.text_input("Color")

        n5, n6, n7 = st.columns(3)
        f_dano    = n5.selectbox("Tipo de daño", ["MEDIANO", "LEVE", "GRAVE"])
        f_ingreso = n6.date_input("Fecha ingreso", value=hoy)
        f_entrega = n7.date_input("Fecha entrega estimada", value=hoy + timedelta(days=14))

        st.markdown("**Datos del cliente**")
        c1, c2, c3, c4 = st.columns(4)
        f_nombre  = c1.text_input("Nombre cliente")
        f_cia     = c2.text_input("Compañía (CIA)")
        f_ase     = c3.text_input("ASE")
        f_tel     = c4.text_input("Teléfono")

        st.markdown("**Asignación de equipo**")
        a1, a2, a3 = st.columns([1, 1, 2])
        auto_asignar = a1.checkbox("🤖 Asignación automática", value=True)

        ots_actual = load_ots()
        if auto_asignar:
            sug_des = mejor_desabollador(ots_actual, f_dano)
            sug_pin = mejor_pintor(ots_actual, f_dano)
            a2.info(f"Se asignará:\n**Des:** {sug_des}\n**Pin:** {sug_pin}")
            f_des = sug_des
            f_pin = sug_pin
        else:
            f_des = a2.selectbox("Desabollador", DESABOLLADORES)
            f_pin = a3.selectbox("Pintor", PINTORES)

        f_comentario = st.text_area("Comentario / observaciones", height=60)
        f_repuestos  = st.text_input("Estado repuestos")

        submitted = st.form_submit_button("🚀 Programar vehículo", type="primary", use_container_width=True)

        if submitted:
            if not f_ot or not f_patente or not f_modelo:
                st.error("OT, Patente y Modelo son obligatorios")
            else:
                ot_id = f"OT-{f_ot.strip()}"
                ots_ids = {o['ot'] for o in ots_actual}
                if ot_id in ots_ids:
                    st.error(f"{ot_id} ya existe en el sistema")
                else:
                    prog = programar_ot(f_des, f_pin, f_dano, f_ingreso, ots_actual)
                    data = {
                        'ot': ot_id, 'nombre': f_nombre, 'cia': f_cia, 'ase': f_ase,
                        'modelo': f_modelo, 'color': f_color, 'patente': f_patente,
                        'dano': f_dano, 'desabollador': f_des, 'pintor': f_pin,
                        'etapa_actual': 'desabolladura',
                        'fecha_ingreso': f_ingreso.strftime('%Y-%m-%d'),
                        'fecha_entrega': f_entrega.strftime('%Y-%m-%d'),
                        'comentario': f_comentario, 'repuestos': f_repuestos, 'telefono': f_tel,
                        'programacion': prog,
                    }
                    save_ot(data)
                    st.success(f"✅ {ot_id} programado correctamente — Desabollador: {f_des} · Pintor: {f_pin}")
                    st.rerun()

# ══════════════════════════════════════════════════════
# TAB 5 — CONFIG
# ══════════════════════════════════════════════════════
with tab_config:
    st.markdown("### ⚙️ Configuración")

    # Panel de pesos de asignación
    with st.expander("⚖️ Pesos de asignación automática", expanded=False):
        st.caption("Define qué porcentaje de OTs se asigna a cada operario según el tipo de daño. El sistema respeta estos pesos al asignar automáticamente.")
        pesos = get_pesos()
        DANOS = ['LEVE', 'MEDIANO', 'GRAVE']

        st.markdown("**Desabolladores**")
        for dano in DANOS:
            st.markdown(f"*{dano}*")
            cols = st.columns(len(DESABOLLADORES))
            total = 0
            nuevos = {}
            for i, des in enumerate(DESABOLLADORES):
                val = cols[i].number_input(
                    des, min_value=0, max_value=100,
                    value=int(pesos['desabolladores'][dano].get(des, 33)),
                    key=f"peso_des_{dano}_{des}", step=5
                )
                nuevos[des] = val
                total += val
            color = "🟢" if total == 100 else "🔴"
            st.caption(f"{color} Total: {total}% {'✓' if total == 100 else '— debe sumar 100%'}")
            pesos['desabolladores'][dano] = nuevos

        st.markdown("---")
        st.markdown("**Pintores**")
        for dano in DANOS:
            st.markdown(f"*{dano}*")
            cols = st.columns(len(PINTORES))
            total = 0
            nuevos = {}
            for i, pin in enumerate(PINTORES):
                val = cols[i].number_input(
                    pin, min_value=0, max_value=100,
                    value=int(pesos['pintores'][dano].get(pin, 50)),
                    key=f"peso_pin_{dano}_{pin}", step=5
                )
                nuevos[pin] = val
                total += val
            color = "🟢" if total == 100 else "🔴"
            st.caption(f"{color} Total: {total}% {'✓' if total == 100 else '— debe sumar 100%'}")
            pesos['pintores'][dano] = nuevos

        st.session_state.pesos_asignacion = pesos
        st.info("💡 Los pesos se aplican en la próxima asignación automática. Se reinician al cerrar la app.")

    st.markdown("---")
    # Importar planilla
    with st.expander("📂 Importar desde planilla Excel", expanded=True):
        st.caption("Sube la planilla del taller para cargar los vehículos en trabajo.")
        archivo = st.file_uploader("Seleccionar archivo Excel", type=["xlsx", "xls"])
        if archivo:
            try:
                df_raw = pd.read_excel(archivo, sheet_name='EN TRABAJO', skiprows=1)
                df_raw.columns = [str(c).strip() for c in df_raw.columns]
                st.success(f"✅ {len(df_raw)} filas encontradas")
                st.dataframe(df_raw[['O T','MODELO','PATENTE','DES','PIN','EST']].head(5),
                           hide_index=True, use_container_width=True)
                if st.button("📥 Importar al sistema", type="primary"):
                    n = importar_planilla(df_raw)
                    st.success(f"✅ {n} vehículos importados")
                    st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")

    # Estadísticas
    with st.expander("📊 Resumen del sistema"):
        total = len(ots)
        por_etapa = Counter(o['etapa_actual'] for o in ots)
        por_des   = Counter(o['desabollador'] for o in ots)
        por_pin   = Counter(o['pintor'] for o in ots)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total activos", total)
        m2.metric("Urgentes (≤3 días)", n_urgentes)
        m3.metric("Listos para entrega", por_etapa.get('listo', 0))
        m4.metric("En desabolladura", por_etapa.get('desabolladura', 0))

        st.markdown("**Por etapa:**")
        for e in ETAPAS:
            n = por_etapa.get(e, 0)
            st.markdown(f"- {ETAPA_LABEL[e]}: **{n}** vehículos")

    # Peligro
    with st.expander("⚠️ Zona de peligro"):
        st.warning("Esta acción elimina TODOS los vehículos del sistema y no se puede deshacer.")
        if 'confirmar_limpiar' not in st.session_state:
            st.session_state.confirmar_limpiar = False
        if not st.session_state.confirmar_limpiar:
            if st.button("🗑️ Limpiar toda la planilla", type="secondary"):
                st.session_state.confirmar_limpiar = True
                st.rerun()
        else:
            st.error("¿Estás seguro? Esta acción no se puede deshacer.")
            c1, c2 = st.columns(2)
            if c1.button("✅ Sí, eliminar todo", type="primary"):
                if supa:
                    supa.table("ots").delete().neq("ot", "").execute()
                else:
                    c = conn.cursor()
                    c.execute("DELETE FROM ots")
                    conn.commit()
                st.session_state.confirmar_limpiar = False
                st.success("Planilla limpiada correctamente")
                st.rerun()
            if c2.button("❌ Cancelar"):
                st.session_state.confirmar_limpiar = False
                st.rerun()
