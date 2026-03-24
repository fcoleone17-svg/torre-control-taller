import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date
import json
from io import BytesIO
import plotly.figure_factory as ff
import plotly.graph_objects as go

# Google Sheets integration
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

# Supabase integration
try:
    from supabase import create_client
    SUPABASE_AVAILABLE = True
except ImportError:
    SUPABASE_AVAILABLE = False

# ── Detectar si estamos en la nube (Streamlit Cloud) o local
def is_cloud():
    try:
        return "supabase_url" in st.secrets
    except:
        return False

# ── Capa de abstraccion de BD
class DB:
    def __init__(self):
        if is_cloud() and SUPABASE_AVAILABLE:
            self.mode = "supabase"
            self.client = create_client(
                st.secrets["supabase_url"],
                st.secrets["supabase_key"]
            )
        else:
            self.mode = "sqlite"
            self.conn = sqlite3.connect("taller_db.db", check_same_thread=False)
            self.cur = self.conn.cursor()
            self._init_sqlite()

    def _init_sqlite(self):
        self.cur.execute('''CREATE TABLE IF NOT EXISTS vehiculos (
            ot TEXT PRIMARY KEY, modelo TEXT, patente TEXT, color TEXT,
            tipo_dano TEXT, asignaciones_json TEXT, fecha_entrega TEXT,
            estado TEXT, repuestos TEXT, nombre TEXT, cia TEXT,
            liquidador TEXT, ase TEXT, loc TEXT, comentario TEXT, fecha_llega TEXT
        )''')
        for col in ['nombre','cia','liquidador','ase','loc','comentario','fecha_llega']:
            try: self.cur.execute(f'ALTER TABLE vehiculos ADD COLUMN {col} TEXT')
            except: pass
        self.conn.commit()

    def execute(self, sql, params=()):
        if self.mode == "sqlite":
            self.cur.execute(sql, params)
        # Supabase usa metodos propios — ver metodos especificos abajo

    def commit(self):
        if self.mode == "sqlite":
            self.conn.commit()

    def fetchall(self):
        if self.mode == "sqlite":
            return self.cur.fetchall()
        return []

    def upsert(self, data):
        if self.mode == "sqlite":
            self.cur.execute(
                "INSERT OR REPLACE INTO vehiculos VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (data['ot'], data['modelo'], data['patente'], data['color'],
                 data['tipo_dano'], data['asignaciones_json'], data['fecha_entrega'],
                 data['estado'], data['repuestos'], data.get('nombre',''),
                 data.get('cia',''), data.get('liquidador',''), data.get('ase',''),
                 data.get('loc',''), data.get('comentario',''), data.get('fecha_llega','')))
            self.conn.commit()
        else:
            self.client.table("vehiculos").upsert(data).execute()

    def update(self, ot, fields):
        if self.mode == "sqlite":
            sets = ", ".join(f"{k}=?" for k in fields.keys())
            vals = list(fields.values()) + [ot]
            self.cur.execute(f"UPDATE vehiculos SET {sets} WHERE ot=?", vals)
            self.conn.commit()
        else:
            self.client.table("vehiculos").update(fields).eq("ot", ot).execute()

    def delete(self, ot):
        if self.mode == "sqlite":
            self.cur.execute("DELETE FROM vehiculos WHERE ot=?", (ot,))
            self.conn.commit()
        else:
            self.client.table("vehiculos").delete().eq("ot", ot).execute()

    def delete_all(self):
        if self.mode == "sqlite":
            self.cur.execute("DELETE FROM vehiculos")
            self.conn.commit()
        else:
            self.client.table("vehiculos").delete().neq("ot", "").execute()

    def select_all(self):
        if self.mode == "sqlite":
            self.cur.execute("SELECT * FROM vehiculos")
            return self.cur.fetchall()
        else:
            resp = self.client.table("vehiculos").select("*").execute()
            # Convertir a lista de tuplas en el mismo orden que SQLite
            cols = ['ot','modelo','patente','color','tipo_dano','asignaciones_json',
                    'fecha_entrega','estado','repuestos','nombre','cia','liquidador',
                    'ase','loc','comentario','fecha_llega']
            rows = []
            for r in resp.data:
                rows.append(tuple(r.get(c, '') or '' for c in cols))
            return rows

db = DB()

# ────────────────────────────────────────────────
# 1. CONFIGURACIÓN Y REGLAS
# ────────────────────────────────────────────────
st.set_page_config(page_title="Torre de Control Taller", layout="wide", page_icon="🗼")

# ── Google Sheets config
SHEET_ID = "1pZtya0aUUMDzq4xVxJ3HsXZb3OOuiEyBVTOKmV3NDqE"
SHEET_GID = "1336876543"
CREDENTIALS_FILE = "torre-control-taller-b8df7b8a6c9c.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly",
          "https://www.googleapis.com/auth/drive.readonly"]

@st.cache_resource
def get_gsheet_client():
    if not GSPREAD_AVAILABLE:
        return None
    try:
        if is_cloud():
            # En la nube usar secrets de Streamlit
            creds = Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]),
                scopes=SCOPES
            )
        else:
            # Local usar archivo JSON
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        return None

def leer_sheet():
    gc = get_gsheet_client()
    if not gc:
        return None
    try:
        sh = gc.open_by_key(SHEET_ID)
        ws = next((w for w in sh.worksheets() if str(w.id) == SHEET_GID), sh.sheet1)
        # Leer todos los valores crudos
        all_values = ws.get_all_values()
        if not all_values:
            return None
        # Buscar la fila que contiene los headers reales (buscar fila con 'O T' o 'PATENTE')
        header_row_idx = 0
        for i, row in enumerate(all_values):
            row_upper = [str(c).strip().upper() for c in row]
            if 'O T' in row_upper or 'PATENTE' in row_upper or 'MODELO' in row_upper:
                header_row_idx = i
                break
        headers = all_values[header_row_idx]
        # Limpiar headers duplicados o vacíos
        seen = {}
        clean_headers = []
        for h in headers:
            h = str(h).strip()
            if h == '':
                h = f'_col_{len(clean_headers)}'
            if h in seen:
                seen[h] += 1
                h = f"{h}_{seen[h]}"
            else:
                seen[h] = 0
            clean_headers.append(h)
        # Construir DataFrame desde la fila siguiente al header
        data_rows = all_values[header_row_idx + 1:]
        df = pd.DataFrame(data_rows, columns=clean_headers)
        # Filtrar filas completamente vacías
        df = df[df.apply(lambda r: any(str(v).strip() for v in r), axis=1)]
        return df
    except Exception as e:
        st.warning(f"No se pudo leer el Sheet: {e}")
        return None


TEAMS = ['Rojas', 'González', 'Carvajal', 'Percy PIN', 'Percy PUL', 'Huaiquifil PIN', 'Huaiquifil PUL']
STAGES_ORDER = ['desabolladura', 'pintura', 'armado', 'pulido', 'terminaciones', 'lavado']

# Capacidad maxima de OTs simultaneas por operario
CAPACIDAD = {
    'Rojas': 1, 'González': 1, 'Carvajal': 1,
    'Percy PIN': 1, 'Percy PUL': 1,
    'Huaiquifil PIN': 2, 'Huaiquifil PUL': 2,
}
# Secuencia real: Desabolladura -> Pintura -> Armado -> Pulido -> Terminaciones -> Lavado
# Mismo equipo hace desabolladura y armado (con pintura en medio)
# Mismo equipo hace pintura y pulido (con armado en medio)

STAGE_TO_GROUP = {
    'desabolladura': 'desabolladura', 'armado': 'armado',
    'pintura': 'pintura', 'pulido': 'pulido',
    'terminaciones': 'terminaciones', 'lavado': 'lavado'
}

ROTATION_ORDER = {
    'desabolladura': ['Rojas', 'González', 'Carvajal'],
    'armado':        ['Rojas', 'González', 'Carvajal'],
    'pintura':       ['Percy PIN', 'Huaiquifil PIN'],
    'pulido':        ['Percy PUL', 'Huaiquifil PUL']
}

DURACIONES = {
    'leve':    {'desabolladura': 1, 'armado': 1, 'pintura': 1, 'pulido': 1, 'terminaciones': 1, 'lavado': 1},
    'mediano': {'desabolladura': 2, 'armado': 1, 'pintura': 2, 'pulido': 1, 'terminaciones': 1, 'lavado': 1},
    'grave':   {'desabolladura': 5, 'armado': 3, 'pintura': 5, 'pulido': 2, 'terminaciones': 2, 'lavado': 1},
}

STAGE_COLORS = {
    'desabolladura': '#E74C3C', 'pintura': '#3498DB', 'armado': '#E67E22',
    'pulido': '#9B59B6', 'terminaciones': '#1ABC9C', 'lavado': '#2ECC71'
}

FECHA_BASE = date(2026, 1, 1)
FERIADOS_2026 = [
    date(2026, 1, 1), date(2026, 4, 3), date(2026, 5, 1),
    date(2026, 9, 18), date(2026, 9, 19), date(2026, 12, 25)
]

ALERTA_DIAS = 3  # días antes de entrega para alertar

# ────────────────────────────────────────────────
# 2. FUNCIONES CORE
# ────────────────────────────────────────────────
def es_dia_laboral(fecha):
    return fecha.weekday() < 5 and fecha not in FERIADOS_2026

def buscar_bloque_libre(team, calendario, auto_disponible_idx, duracion):
    hoy_idx = (date.today() - FECHA_BASE).days
    inicio_busqueda = max(auto_disponible_idx, hoy_idx)
    cap = CAPACIDAD.get(team, 1) if team else 999

    while True:
        posible_inicio = inicio_busqueda
        dias_encontrados = 0
        temp_idx = inicio_busqueda
        while dias_encontrados < duracion:
            fecha_check = FECHA_BASE + timedelta(days=temp_idx)
            if es_dia_laboral(fecha_check):
                ocupacion = calendario.get(temp_idx, 0) if hasattr(calendario, 'get') else (1 if temp_idx in calendario else 0)
                if team is not None and ocupacion >= cap:
                    break
                dias_encontrados += 1
            temp_idx += 1
        if dias_encontrados == duracion:
            return posible_inicio, temp_idx - 1
        inicio_busqueda += 1

def pul_de(pin_team):
    # Dado un equipo PIN, devuelve su contraparte PUL
    if 'PIN' in str(pin_team):
        return pin_team.replace('PIN', 'PUL')
    if pin_team == 'Percy':
        return 'Percy PUL'
    if pin_team == 'Huaiquifil':
        return 'Huaiquifil PUL'
    return pin_team

def elegir_equipo_optimo(grupo, auto_disponible_idx, duracion, calendarios):
    """Elige automáticamente el equipo que termina antes."""
    if grupo not in ROTATION_ORDER:
        return None
    mejor_team = None
    mejor_fin = float('inf')
    for team in ROTATION_ORDER[grupo]:
        _, fin = buscar_bloque_libre(team, calendarios.get(team, set()), auto_disponible_idx, duracion)
        if fin < mejor_fin:
            mejor_fin = fin
            mejor_team = team
    return mejor_team

def verificar_eficiencia(equipo_elegido, grupo, auto_disponible_idx, duracion, calendarios):
    if not equipo_elegido or grupo not in ROTATION_ORDER:
        return True
    _, fin_elegido = buscar_bloque_libre(equipo_elegido, calendarios.get(equipo_elegido, set()), auto_disponible_idx, duracion)
    for otro in ROTATION_ORDER[grupo]:
        if otro == equipo_elegido:
            continue
        _, fin_otro = buscar_bloque_libre(otro, calendarios.get(otro, set()), auto_disponible_idx, duracion)
        if fin_otro < fin_elegido:
            return False
    return True

def calcular_carga_proximos_dias(team, calendarios, dias=14):
    hoy_idx = (date.today() - FECHA_BASE).days
    cal = calendarios.get(team, {})
    return sum(1 for i in range(hoy_idx, hoy_idx + dias) if cal.get(i, 0) > 0)

# ────────────────────────────────────────────────
# 3. DB Y PERSISTENCIA
# ────────────────────────────────────────────────
# BD inicializada en clase DB arriba

def load_vehicles():
    vehicles = []
    for row in db.select_all():
        try:
            asig = json.loads(row[5])
            for s in asig:
                asig[s]['start'] = datetime.strptime(asig[s]['start'], '%Y-%m-%d').date()
                asig[s]['end'] = datetime.strptime(asig[s]['end'], '%Y-%m-%d').date()
            estado_v = row[7] if len(row) > 7 else 'Programado'
            asig = corregir_pul(asig, estado=estado_v)
            vehicles.append({
                'OT': row[0], 'Modelo': row[1], 'Patente': row[2], 'Color': row[3],
                'Tipo Daño': row[4], 'Asignaciones': asig,
                'Fecha Entrega': datetime.strptime(row[6], '%Y-%m-%d').date(),
                'Estado': row[7], 'Repuestos': row[8],
                'Nombre': row[9] if len(row) > 9 else '',
                'CIA': row[10] if len(row) > 10 else '',
                'Liquidador': row[11] if len(row) > 11 else '',
                'ASE': row[12] if len(row) > 12 else '',
                'LOC': row[13] if len(row) > 13 else '',
                'Comentario': row[14] if len(row) > 14 else '',
                'Fecha Llega': row[15] if len(row) > 15 else '',
            })
        except:
            continue
    return vehicles

def normalizar_equipo_pintura(team):
    # Convierte nombres viejos de pintura a formato PIN
    if team == 'Percy':        return 'Percy PIN'
    if team == 'Huaiquifil':   return 'Huaiquifil PIN'
    if team == 'Percy PUL':    return 'Percy PIN'
    if team == 'Huaiquifil PUL': return 'Huaiquifil PIN'
    return team

# Mapa de etapas completadas según estado actual
# El estado actual está EN CURSO, solo las etapas ANTERIORES son completadas
ETAPAS_COMPLETADAS_POR_ESTADO = {
    'En Desabolladura': [],
    'En Pintura':       ['desabolladura'],
    'En Armado':        ['desabolladura', 'pintura'],
    'En Pulido':        ['desabolladura', 'pintura', 'armado'],
    'En Calidad':       ['desabolladura', 'pintura', 'armado', 'pulido'],
    'Lavado':           ['desabolladura', 'pintura', 'armado', 'pulido', 'terminaciones'],
    'Listo para Entrega': ['desabolladura', 'pintura', 'armado', 'pulido', 'terminaciones', 'lavado'],
    'Programado':       [],
}

def corregir_pul(asig, estado=None):
    # 1. Normaliza nombres de equipos pintura/pulido
    pin_info = asig.get('pintura', {})
    pul_info = asig.get('pulido', {})

    if pin_info and not pin_info.get('completado'):
        pin_info['team'] = normalizar_equipo_pintura(pin_info.get('team', ''))

    pin_team = pin_info.get('team', '')
    if pul_info and not pul_info.get('completado') and pin_team:
        pul_info['team'] = pul_de(pin_team)

    # 2. Si se pasa el estado, marcar etapas anteriores como completadas
    if estado:
        etapas_listas = ETAPAS_COMPLETADAS_POR_ESTADO.get(estado, [])
        for s in etapas_listas:
            if s in asig and not asig[s].get('completado'):
                asig[s]['completado'] = True

    return asig

def rebuild_calendars(vehicles):
    from collections import Counter
    new_cals = {team: Counter() for team in TEAMS}
    for v in vehicles:
        if v['Estado'] == "Listo para Entrega":
            continue
        for info in v['Asignaciones'].values():
            if info.get('completado'):
                continue
            t = info.get('team')
            if t in TEAMS:
                s_idx = (info['start'] - FECHA_BASE).days
                e_idx = (info['end'] - FECHA_BASE).days
                for i in range(s_idx, e_idx + 1):
                    if es_dia_laboral(FECHA_BASE + timedelta(days=i)):
                        new_cals[t][i] += 1
    return new_cals

def avanzar_estado_automatico(v, hoy_auto):
    estado_actual = v['Estado']
    if estado_actual in ('Listo para Entrega', 'Entregado'):
        return estado_actual
    asigs = v['Asignaciones']
    todas_pasaron = all(
        info.get('completado', False) or (info.get('end') and info['end'] < hoy_auto)
        for info in asigs.values()
    )
    if todas_pasaron:
        return 'Listo para Entrega'
    secuencia = ['desabolladura', 'pintura', 'armado', 'pulido', 'terminaciones', 'lavado']
    estado_map = {
        'desabolladura': 'En Pintura',
        'pintura':       'En Armado',
        'armado':        'En Pulido',
        'pulido':        'En Calidad',
        'terminaciones': 'Lavado',
        'lavado':        'Listo para Entrega',
    }
    nuevo_estado = estado_actual
    for etapa in secuencia:
        info = asigs.get(etapa, {})
        ya = info.get('completado', False) or (info.get('end') and info['end'] < hoy_auto)
        if ya:
            nuevo_estado = estado_map.get(etapa, estado_actual)
        else:
            break
    return nuevo_estado

def persistir_correcciones(vehicles):
    hoy_auto = date.today()
    for v in vehicles:
        asigs = v['Asignaciones']
        nuevo_estado = avanzar_estado_automatico(v, hoy_auto)
        asigs_db = {
            k: {**info,
                'start': info['start'].strftime('%Y-%m-%d'),
                'end':   info['end'].strftime('%Y-%m-%d')}
            for k, info in asigs.items()
        }
        db.update(v['OT'], {
            'estado': nuevo_estado,
            'asignaciones_json': json.dumps(asigs_db)
        })


if 'vehicles' not in st.session_state:
    st.session_state.vehicles = load_vehicles()
    persistir_correcciones(st.session_state.vehicles)
    st.session_state.calendarios = rebuild_calendars(st.session_state.vehicles)
    st.session_state.sheet_sync_done = False

# Auto-sync desde Google Sheet al abrir la app (una vez por sesión)
if not st.session_state.get('sheet_sync_done', False) and GSPREAD_AVAILABLE:
    with st.spinner("🔄 Sincronizando con Google Sheets..."):
        df_sheet = leer_sheet()
        if df_sheet is not None:
            st.session_state.sheet_sync_done = True
            st.session_state.sheet_data = df_sheet
            st.toast("✅ Google Sheets sincronizado", icon="📊")
        else:
            st.session_state.sheet_sync_done = True

# ────────────────────────────────────────────────
# 4. ALERTAS DE VENCIMIENTO (banner superior)
# ────────────────────────────────────────────────
hoy = date.today()
alertas = []
for v in st.session_state.vehicles:
    if v['Estado'] == "Listo para Entrega":
        continue
    dias_restantes = (v['Fecha Entrega'] - hoy).days
    if dias_restantes < 0:
        alertas.append(f"🚨 **{v['OT']} - {v['Modelo']} ({v['Patente']})** lleva **{abs(dias_restantes)} días de atraso**")
    elif dias_restantes <= ALERTA_DIAS:
        alertas.append(f"⚠️ **{v['OT']} - {v['Modelo']} ({v['Patente']})** vence en **{dias_restantes} día(s)**")




# ────────────────────────────────────────────────
# 6. PANEL CENTRAL
# ────────────────────────────────────────────────
st.markdown("## 🗼 Torre de Control — Taller")

# TABS principales
# Badge con número de alertas en el título
n_alertas = len(alertas)
alerta_label = f"🔔 Alertas ({n_alertas})" if n_alertas > 0 else "🔔 Alertas"
tab_tablero, tab_gantt, tab_calendario, tab_carga, tab_alertas = st.tabs([
    "📋 Tablero", "📊 Gantt", "📅 Calendario Semanal", "⚙️ Gestión", alerta_label
])

# ── TAB 1: TABLERO ──
with tab_tablero:
    if st.session_state.vehicles:

        if 'editar_ot' not in st.session_state:
            st.session_state.editar_ot = None

        ESTADOS_LIST = ["Programado", "En Desabolladura", "En Pintura", "En Armado",
                        "En Pulido", "En Calidad", "Lavado", "Listo para Entrega"]

        # Colores de fila por estado
        COLOR_ESTADO = {
            "Programado":          "#F8F9FA",
            "En Desabolladura":    "#FFF3CD",
            "En Pintura":          "#D0E8FF",
            "En Armado":           "#FFE5CC",
            "En Pulido":           "#EAD6F5",
            "En Calidad":          "#D6EFD8",
            "Lavado":              "#D6EFD8",
            "Listo para Entrega":  "#C8F0C8",
        }

        # ── Filtros rápidos
        with st.container():
            fc1, fc2, fc3 = st.columns([2, 2, 3])
            todos_estados = ["Todos"] + ESTADOS_LIST
            filtro_estado = fc1.selectbox("🔍 Filtrar por estado",
                                           todos_estados, key="filtro_estado")
            todos_operarios = ["Todos"] + TEAMS
            filtro_operario = fc2.selectbox("🔍 Filtrar por operario",
                                             todos_operarios, key="filtro_operario")
            filtro_texto = fc3.text_input("🔍 Buscar patente o modelo", key="filtro_texto")

        # ── Construir y filtrar vehículos
        vehicles_filtrados = st.session_state.vehicles
        if filtro_estado != "Todos":
            vehicles_filtrados = [v for v in vehicles_filtrados if v['Estado'] == filtro_estado]
        if filtro_operario != "Todos":
            vehicles_filtrados = [v for v in vehicles_filtrados
                                  if any(info.get('team') == filtro_operario
                                         for info in v['Asignaciones'].values())]
        if filtro_texto:
            txt = filtro_texto.lower()
            vehicles_filtrados = [v for v in vehicles_filtrados
                                  if txt in v['Patente'].lower() or txt in v['Modelo'].lower()]

        st.caption(f"Mostrando {len(vehicles_filtrados)} de {len(st.session_state.vehicles)} OTs")

        # ── Construir filas
        rows = []
        for v in vehicles_filtrados:
            etapas_ok = set(ETAPAS_COMPLETADAS_POR_ESTADO.get(v['Estado'], []))
            row = {'OT': v['OT'], 'Patente': v['Patente'],
                   'Modelo': v['Modelo'], 'Color': v['Color']}
            for s in STAGES_ORDER:
                info = v['Asignaciones'].get(s, {})
                es_comp = (s in etapas_ok or info.get('completado', False)
                           or (info.get('end') and info['end'] < hoy))
                if es_comp:
                    t_name = str(info.get('team','')) if info.get('team') not in ('Completado', None, '') else ''
                    sufijo = f" ({t_name})" if t_name and t_name != 'Completado' else ''
                    row[s.capitalize()] = f"✅{sufijo}"
                elif info:
                    t_name = str(info.get('team','')) if info.get('team') else "—"
                    inicio = info['start'].strftime('%d/%m') if info.get('start') else "—"
                    row[s.capitalize()] = f"{t_name} ({inicio})"
                else:
                    row[s.capitalize()] = "—"
            dias_r = (v['Fecha Entrega'] - hoy).days
            if dias_r < 0:
                row['Entrega'] = f"🚨 {v['Fecha Entrega'].strftime('%d/%m/%y')}"
            elif dias_r <= ALERTA_DIAS:
                row['Entrega'] = f"⚠️ {v['Fecha Entrega'].strftime('%d/%m/%y')}"
            else:
                row['Entrega'] = v['Fecha Entrega'].strftime('%d/%m/%y')
            row['Entrega_sort'] = v['Fecha Entrega']
            row['_estado'] = v['Estado']
            row['Estado'] = v['Estado']
            row['Repuestos'] = v['Repuestos']
            row['Cliente'] = v.get('Nombre', '') or ''
            row['CIA'] = v.get('CIA', '') or ''
            rows.append(row)

        df_v = pd.DataFrame(rows).sort_values('Entrega_sort')
        df_display = df_v.drop(columns=['Entrega_sort', '_estado'])

        # ── Estilos: color de fila por estado + colores de celda
        def style_row(row):
            estado = df_v.loc[row.name, '_estado'] if row.name in df_v.index else ''
            bg = COLOR_ESTADO.get(estado, '#FFFFFF')
            styles = [f'background-color: {bg}'] * len(row)
            for i, val in enumerate(row):
                if isinstance(val, str) and val.startswith('✅'):
                    styles[i] = 'color:#1a7a1a; font-weight:bold; background-color:#e6f9e6'
            return styles

        st.dataframe(
            df_display.style.apply(style_row, axis=1),
            use_container_width=True, hide_index=True, height=440
        )

        # ── Edición inline debajo del tablero
        st.markdown("---")
        ed1, ed2, ed3 = st.columns([3, 1, 2])
        ot_sel_edit = ed1.selectbox(
            "✏️ Seleccionar OT para actualizar:",
            [v['OT'] for v in st.session_state.vehicles],
            key="sel_editar_ot")
        if ed2.button("✏️ Editar", type="primary", key="btn_abrir_editor"):
            st.session_state.editar_ot = (ot_sel_edit
                if st.session_state.editar_ot != ot_sel_edit else None)
            st.rerun()
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_display.to_excel(writer, index=False)
        ed3.download_button("📥 Descargar Excel", output.getvalue(),
                             "Planificacion_Taller.xlsx", use_container_width=True)

        # ── Panel de edición
        if st.session_state.editar_ot:
            veh_edit = next((v for v in st.session_state.vehicles
                             if v['OT'] == st.session_state.editar_ot), None)
            if veh_edit:
                ot_id = veh_edit['OT']
                color_panel = COLOR_ESTADO.get(veh_edit['Estado'], '#f0f7ff')
                st.markdown(
                    f"<div style='background:{color_panel};padding:16px;"
                    f"border-radius:10px;border:1px solid #90c0e8;margin-top:8px'>"
                    f"<b>✏️ {ot_id} — {veh_edit['Patente']} &nbsp; {veh_edit['Modelo']} "
                    f"&nbsp; {veh_edit['Color']}</b></div>",
                    unsafe_allow_html=True)
                st.markdown("")

                p1, p2 = st.columns(2)
                est_idx = ESTADOS_LIST.index(veh_edit['Estado']) if veh_edit['Estado'] in ESTADOS_LIST else 0
                nuevo_estado = p1.selectbox("Estado", ESTADOS_LIST,
                                             index=est_idx, key=f"est_{ot_id}")
                rep_list = ["Pendiente ⏳", "Recibidos Parcial 📦", "Completos ✅"]
                rep_idx = rep_list.index(veh_edit['Repuestos']) if veh_edit['Repuestos'] in rep_list else 0
                nuevo_rep = p2.selectbox("Repuestos", rep_list,
                                          index=rep_idx, key=f"rep_{ot_id}")

                # Checkboxes automáticos según estado seleccionado
                etapas_auto = set(ETAPAS_COMPLETADAS_POR_ESTADO.get(nuevo_estado, []))
                st.markdown("**Etapas completadas** — se marcan automáticamente según el estado:")
                chk_cols = st.columns(6)
                nuevos_comp = {}
                for i, s in enumerate(STAGES_ORDER):
                    info = veh_edit['Asignaciones'].get(s, {})
                    # Auto-marcar según estado nuevo O si ya estaba completada antes
                    ya = (s in etapas_auto
                          or info.get('completado', False)
                          or (info.get('end') and info['end'] < hoy))
                    nuevos_comp[s] = chk_cols[i].checkbox(
                        s.capitalize(), value=ya, key=f"chk_{ot_id}_{s}")

                c_guardar, c_entregar, c_cerrar = st.columns([2, 2, 1])
                if c_guardar.button("💾 Guardar", key=f"save_{ot_id}", type="primary"):
                    asigs = veh_edit['Asignaciones']
                    for s in STAGES_ORDER:
                        if s in asigs:
                            asigs[s]['completado'] = nuevos_comp[s]
                    asigs_db = {k: {**info,
                                    'start': info['start'].strftime('%Y-%m-%d'),
                                    'end':   info['end'].strftime('%Y-%m-%d')}
                                for k, info in asigs.items()}
                    db.update(ot_id, {
                        'estado': nuevo_estado,
                        'repuestos': nuevo_rep,
                        'asignaciones_json': json.dumps(asigs_db)
                    })
                    st.session_state.vehicles = load_vehicles()
                    st.session_state.calendarios = rebuild_calendars(st.session_state.vehicles)
                    st.session_state.editar_ot = None
                    st.success(f"✅ {ot_id} actualizada correctamente")
                    st.rerun()

                if c_entregar.button("🚗 Marcar como Entregado", key=f"entregar_{ot_id}"):
                    db.delete(ot_id)
                    st.session_state.vehicles = load_vehicles()
                    st.session_state.calendarios = rebuild_calendars(st.session_state.vehicles)
                    st.session_state.editar_ot = None
                    st.success(f"🚗 {ot_id} entregada y retirada del tablero")
                    st.rerun()

                if c_cerrar.button("✖", key=f"close_{ot_id}", help="Cerrar"):
                    st.session_state.editar_ot = None
                    st.rerun()
    else:
        st.info("No hay vehículos programados. Agrega una OT desde el panel inferior.")

# ── TAB 2: GANTT ──
with tab_gantt:
    st.subheader("📊 Diagrama de Gantt — Línea de tiempo por OT")

    if st.session_state.vehicles:
        gantt_data = []
        for v in st.session_state.vehicles:
            for s in STAGES_ORDER:
                info = v['Asignaciones'].get(s)
                if info:
                    gantt_data.append({
                        'Task': f"{v['OT']} — {v['Patente']}",
                        'Start': datetime.combine(info['start'], datetime.min.time()),
                        'Finish': datetime.combine(info['end'] + timedelta(days=1), datetime.min.time()),
                        'Resource': s.capitalize(),
                        'Operario': info.get('team', 'N/A'),
                        'Eficiente': info.get('eficiente', True)
                    })

        if gantt_data:
            df_gantt = pd.DataFrame(gantt_data)

            # Filtro por operario
            col_f1, col_f2 = st.columns(2)
            filtro_op = col_f1.multiselect("Filtrar por operario:", TEAMS, default=TEAMS)
            filtro_etapa = col_f2.multiselect("Filtrar por etapa:", [s.capitalize() for s in STAGES_ORDER],
                                               default=[s.capitalize() for s in STAGES_ORDER])

            df_filtrado = df_gantt[
                df_gantt['Operario'].isin(filtro_op) &
                df_gantt['Resource'].isin(filtro_etapa)
            ]

            if not df_filtrado.empty:
                # Solo incluir colores de etapas presentes en los datos filtrados
                etapas_presentes = df_filtrado['Resource'].unique()
                colores_filtrados = {e: STAGE_COLORS.get(e.lower(), '#95A5A6') for e in etapas_presentes}

                fig = ff.create_gantt(
                    df_filtrado[['Task', 'Start', 'Finish', 'Resource']],
                    colors=colores_filtrados,
                    index_col='Resource',
                    show_colorbar=True,
                    group_tasks=True,
                    showgrid_x=True,
                    showgrid_y=True,
                    title='Programación del Taller'
                )

                # Marcar ineficiencias con punto rojo
                for _, row in df_filtrado[~df_filtrado['Eficiente']].iterrows():
                    fig.add_shape(type="rect",
                        x0=row['Start'], x1=row['Finish'],
                        y0=-0.4, y1=0.4,
                        line=dict(color="red", width=2),
                        fillcolor="rgba(255,0,0,0.1)"
                    )

                fig.update_layout(height=max(400, len(df_filtrado['Task'].unique()) * 40 + 100),
                                  xaxis_title="Fecha", font=dict(size=11))
                st.plotly_chart(fig, use_container_width=True)
                st.caption("🔴 Bordes rojos = asignación no óptima (existe un operario disponible antes)")
            else:
                st.info("No hay datos para los filtros seleccionados.")
    else:
        st.info("Agrega vehículos para ver el Gantt.")

# ── TAB 3: CALENDARIO SEMANAL ──
with tab_calendario:
    st.subheader("📅 Vista de Calendario Semanal")

    # Selector de semana
    col_s1, col_s2 = st.columns([1, 3])
    semana_offset = col_s1.number_input("Semana (0 = actual)", min_value=-4, max_value=12, value=0)
    lunes = hoy - timedelta(days=hoy.weekday()) + timedelta(weeks=semana_offset)
    dias_semana = [lunes + timedelta(days=i) for i in range(5)]  # Lunes a Viernes

    col_s2.markdown(f"**Semana del {lunes.strftime('%d/%m/%Y')} al {dias_semana[-1].strftime('%d/%m/%Y')}**")

    # Tabla de calendario
    # cal_data: {team: {dia: [(ot, etapas_activas)]}}
    cal_data = {t: {d.strftime('%a %d/%m'): {} for d in dias_semana} for t in TEAMS}

    for v in st.session_state.vehicles:
        if v['Estado'] == "Listo para Entrega":
            continue
        etapas_ok = set(ETAPAS_COMPLETADAS_POR_ESTADO.get(v['Estado'], []))
        for etapa, info in v['Asignaciones'].items():
            team = info.get('team')
            if team not in TEAMS:
                continue
            # No mostrar etapas ya completadas en el calendario
            es_comp = (etapa in etapas_ok
                       or info.get('completado', False)
                       or (info.get('end') and info['end'] < hoy))
            if es_comp:
                continue
            for dia in dias_semana:
                if info['start'] <= dia <= info['end'] and es_dia_laboral(dia):
                    key = dia.strftime('%a %d/%m')
                    ot = v['OT']
                    # Agrupar etapas por OT: DES+ARM -> "DES→ARM"
                    if ot not in cal_data[team][key]:
                        cal_data[team][key][ot] = []
                    cal_data[team][key][ot].append(etapa[:3].upper())

    # Construir etiquetas agrupadas por OT
    def etiqueta_ot(ot, etapas):
        etapas_str = "→".join(etapas) if len(etapas) > 1 else etapas[0]
        return f"{ot} ({etapas_str})"

    # Renderizar como tabla HTML con colores
    dias_cols = [d.strftime('%a %d/%m') for d in dias_semana]
    html = "<table style='width:100%; border-collapse:collapse; font-size:13px;'>"
    html += "<tr style='background:#2C3E50; color:white;'><th style='padding:8px;'>Operario</th>"
    for d in dias_cols:
        html += f"<th style='padding:8px; text-align:center;'>{d}</th>"
    html += "</tr>"

    colors_row = ['#ECF0F1', '#FDFEFE']
    for i, team in enumerate(TEAMS):
        bg = colors_row[i % 2]
        html += f"<tr style='background:{bg};'><td style='padding:8px; font-weight:bold;'>{team}</td>"
        for d in dias_cols:
            ots_dia = cal_data[team][d]  # dict {ot: [etapas]}
            tareas = [etiqueta_ot(ot, etapas) for ot, etapas in ots_dia.items()]
            cap = CAPACIDAD.get(team, 1)
            if tareas:
                contenido = "<br>".join([
                    f"<span style='background:#3498DB;color:white;padding:2px 5px;"
                    f"border-radius:3px;font-size:11px;'>{t}</span>"
                    for t in tareas])
                # Rojo si supera capacidad, naranja si al capacidad, azul claro si normal
                if len(tareas) > cap:
                    carga_bg = "#FADBD8"
                elif len(tareas) == cap:
                    carga_bg = "#FEF9E7"
                else:
                    carga_bg = "#EBF5FB"
            else:
                contenido = "<span style='color:#BDC3C7;'>—</span>"
                carga_bg = bg
            html += f"<td style='padding:8px; text-align:center; background:{carga_bg};'>{contenido}</td>"
        html += "</tr>"
    html += "</table>"

    st.markdown(html, unsafe_allow_html=True)
    st.caption("🔵 fondo claro = 1 OT | 🟡 fondo amarillo = al límite de capacidad | 🔴 fondo rojo = sobre capacidad")

    # Gráfico de carga
    st.markdown("---")
    st.subheader("📈 Carga por Operario (próximos 30 días)")
    col_g1, col_g2 = st.columns(2)

    with col_g1:
        hoy_idx_c = (date.today() - FECHA_BASE).days
        carga_datos_list = []
        for t in TEAMS:
            cal_t = st.session_state.calendarios.get(t, {})
            cap_t = CAPACIDAD.get(t, 1)
            dias_hab = sum(1 for i in range(hoy_idx_c, hoy_idx_c + 14)
                           if es_dia_laboral(FECHA_BASE + timedelta(days=i)))
            ocupacion_t = sum(cal_t.get(i, 0) for i in range(hoy_idx_c, hoy_idx_c + 14))
            max_pos = cap_t * dias_hab
            pct_ocup = round(ocupacion_t / max_pos * 100) if max_pos > 0 else 0
            carga_datos_list.append({"Operario": t, "% Ocupación (14d)": pct_ocup})
        carga_datos = pd.DataFrame(carga_datos_list)
        fig_bar = go.Figure(go.Bar(
            x=carga_datos['Operario'], y=carga_datos['% Ocupación (14d)'],
            marker_color=['#E74C3C' if d > 80 else '#F39C12' if d > 40 else '#2ECC71'
                          for d in carga_datos['% Ocupación (14d)']],
            text=[f"{d}%" for d in carga_datos['% Ocupación (14d)']], textposition='outside'
        ))
        fig_bar.update_layout(title="Carga próximos 14 días", yaxis_title="% Ocupación",
                              height=300, showlegend=False)
        st.plotly_chart(fig_bar, use_container_width=True)

    with col_g2:
        st.subheader("🔍 Detalle por Operario")
        op_sel = st.selectbox("Ver tareas de:", TEAMS)
        detalles = []
        for v in st.session_state.vehicles:
            for etapa, info in v['Asignaciones'].items():
                if info.get('team') == op_sel:
                    dias = sum(1 for i in range(
                        (info['start'] - FECHA_BASE).days,
                        (info['end'] - FECHA_BASE).days + 1
                    ) if es_dia_laboral(FECHA_BASE + timedelta(days=i)))
                    detalles.append({
                        "OT": v['OT'], "Etapa": etapa.capitalize(),
                        "Días": dias, "Inicio": info['start'].strftime('%d/%m'),
                        "Termina": info['end'].strftime('%d/%m')
                    })
        if detalles:
            st.dataframe(pd.DataFrame(detalles), hide_index=True, use_container_width=True)
        else:
            st.info(f"Sin tareas para {op_sel}")

# ── TAB 4: GESTIÓN ──
with tab_carga:
    st.subheader("🛠️ Gestión de OT (Repuestos, Estados y Re-asignación)")

    if st.session_state.vehicles:
        v_sel = st.selectbox("Seleccionar OT para modificar", [v['OT'] for v in st.session_state.vehicles])
        veh = next(v for v in st.session_state.vehicles if v['OT'] == v_sel)

        c1, c2, c3, c4 = st.columns(4)
        estados_list = ["Programado", "En Desabolladura", "En Pintura", "En Armado",
                        "En Pulido", "En Calidad", "Lavado", "Listo para Entrega"]
        n_est = c1.selectbox("Estado", estados_list, index=estados_list.index(veh['Estado']))
        rep_list = ["Pendiente ⏳", "Recibidos Parcial 📦", "Completos ✅"]
        n_rep = c2.selectbox("Repuestos", rep_list, index=rep_list.index(veh['Repuestos']))
        n_des = c3.selectbox("Mover Desab./Arm.", ROTATION_ORDER['desabolladura'])
        n_pin = c4.selectbox("Mover Pintura", ROTATION_ORDER['pintura'])

        with st.expander("🏁 Marcar etapas como completadas"):
            cols_etapas = st.columns(3)
            completados_sel = {}
            for i, s in enumerate(STAGES_ORDER):
                info = veh['Asignaciones'].get(s, {})
                ya_completado = info.get('completado', False)
                completados_sel[s] = cols_etapas[i % 3].checkbox(
                    s.capitalize(), value=ya_completado, key=f"comp_{s}"
                )

        with st.expander("✏️ Editar datos del cliente"):
            g1, g2, g3 = st.columns(3)
            n_nombre = g1.text_input("Nombre cliente", value=veh.get('Nombre','') or '')
            n_cia    = g2.text_input("CIA", value=veh.get('CIA','') or '')
            n_liq    = g3.text_input("Liquidador", value=veh.get('Liquidador','') or '')
            g4, g5, g6 = st.columns(3)
            n_ase    = g4.text_input("ASE", value=veh.get('ASE','') or '')
            n_loc    = g5.text_input("LOC", value=veh.get('LOC','') or '')
            try:
                llega_val = datetime.strptime(str(veh.get('Fecha Llega','')), '%Y-%m-%d').date()
            except:
                llega_val = date.today()
            n_llega  = g6.date_input("Fecha llega", value=llega_val)
            n_com    = st.text_area("Comentario", value=veh.get('Comentario','') or '', height=60)

        if st.button("💾 Guardar Cambios", type="primary"):
            asigs = veh['Asignaciones']
            st.session_state.vehicles = [v for v in st.session_state.vehicles if v['OT'] != v_sel]
            temp_cals = rebuild_calendars(st.session_state.vehicles)

            # Aplicar flags de completado
            for s in STAGES_ORDER:
                if s in asigs:
                    asigs[s]['completado'] = completados_sel.get(s, False)
                    if completados_sel.get(s, False):
                        asigs[s]['team'] = 'Completado'

            # Derivar equipo pulido desde pintura elegida
            n_pul = pul_de(n_pin)

            # Solo actualizar equipos en etapas no completadas
            if not asigs.get('desabolladura', {}).get('completado'):
                asigs['desabolladura']['team'] = n_des
            if not asigs.get('armado', {}).get('completado'):
                asigs['armado']['team'] = n_des
            if not asigs.get('pintura', {}).get('completado'):
                asigs['pintura']['team'] = n_pin
            if not asigs.get('pulido', {}).get('completado'):
                asigs['pulido']['team'] = n_pul

            asigs['desabolladura']['eficiente'] = verificar_eficiencia(
                n_des, 'desabolladura',
                (asigs['desabolladura']['start'] - FECHA_BASE).days,
                DURACIONES[veh['Tipo Daño']]['desabolladura'], temp_cals
            )
            asigs['pintura']['eficiente'] = verificar_eficiencia(
                n_pin, 'pintura',
                (asigs['pintura']['start'] - FECHA_BASE).days,
                DURACIONES[veh['Tipo Daño']]['pintura'], temp_cals
            )

            asigs_db = {
                k: {**v, 'start': v['start'].strftime('%Y-%m-%d'), 'end': v['end'].strftime('%Y-%m-%d')}
                for k, v in asigs.items()
            }
            db.update(v_sel, {
                'estado': n_est, 'repuestos': n_rep,
                'asignaciones_json': json.dumps(asigs_db),
                'nombre': n_nombre, 'cia': n_cia, 'liquidador': n_liq,
                'ase': n_ase, 'loc': n_loc, 'comentario': n_com,
                'fecha_llega': n_llega.strftime('%Y-%m-%d')
            })
            st.session_state.vehicles = load_vehicles()
            st.session_state.calendarios = rebuild_calendars(st.session_state.vehicles)
            st.success("✅ Cambios guardados")
            st.rerun()

        st.markdown("---")
        col_del, _ = st.columns([1, 3])
        if col_del.button("🗑️ Eliminar esta OT", type="secondary"):
            db.delete(v_sel)
            st.session_state.vehicles = load_vehicles()
            st.session_state.calendarios = rebuild_calendars(st.session_state.vehicles)
            st.success(f"{v_sel} eliminada")
            st.rerun()

    st.markdown("---")
    st.subheader("📂 Importar datos")
    st.caption("Importa desde Google Sheets o sube un Excel manualmente")

    # Mapeo de abreviaciones a nombres completos
    DES_MAP = {'GO': 'González', 'RO': 'Rojas', 'CA': 'Carvajal',
               'MEC': None, 'EXT': None}  # None = sin asignar
    PIN_MAP = {'PE': 'Percy PIN', 'HU': 'Huaiquifil PIN', 'AS': 'Percy PIN'}
    ESTADO_MAP = {
        'OK.': 'Listo para Entrega', 'LAV': 'Lavado', 'PUL': 'En Pulido',
        'ARM': 'En Armado', 'PIN': 'En Pintura', 'DES': 'En Desabolladura'
    }

    # ── Importar desde Google Sheets
    if GSPREAD_AVAILABLE:
        gs1, gs2 = st.columns([2, 1])
        gs1.markdown("**📊 Google Sheets** — Planilla del taller conectada")
        if gs2.button("🔄 Importar desde Sheets", key="btn_import_sheets"):
            df_sheet = leer_sheet()
            if df_sheet is not None:
                st.session_state.sheet_data = df_sheet
                st.success(f"✅ {len(df_sheet)} filas leídas desde Google Sheets")
            else:
                st.error("No se pudo conectar con Google Sheets")

        if 'sheet_data' in st.session_state and st.session_state.sheet_data is not None:
            df_imp_gs = st.session_state.sheet_data.copy()
            df_imp_gs.columns = [str(c).strip() for c in df_imp_gs.columns]
            st.write(f"**{len(df_imp_gs)} filas disponibles.** Vista previa:")
            st.dataframe(df_imp_gs.head(3), use_container_width=True, hide_index=True)
            if st.button("📥 Cargar al sistema", type="primary", key="btn_cargar_gs"):
                importados_gs = 0
                errores_gs = []
                temp_cals_gs = rebuild_calendars(st.session_state.vehicles)
                ots_existentes = {v['OT'] for v in st.session_state.vehicles}

                # Recopilar OTs válidas del Sheet
                ots_en_sheet = set()
                for _, row in df_imp_gs.iterrows():
                    ot_raw = str(row.get('O T', '')).strip()
                    if ot_raw and ot_raw.lower() not in ('nan', ''):
                        ots_en_sheet.add(f"OT-{ot_raw}")

                # Eliminar OTs que ya no están en el Sheet
                eliminadas_gs = 0
                for ot_app in list(ots_existentes):
                    if ot_app not in ots_en_sheet:
                        db.delete(ot_app)
                        eliminadas_gs += 1
                if eliminadas_gs > 0:
                    pass  # ya commiteado por db.delete
                    # Reconstruir calendarios sin las OTs eliminadas
                    temp_cals_gs = rebuild_calendars(load_vehicles())

                def get_val(row, *keys):
                    # Leer valor de fila limpiando NaN y espacios
                    for k in keys:
                        v = row.get(k, None)
                        if v is not None and str(v).strip().lower() not in ('nan', ''):
                            return str(v).strip()
                    return ''

                for _, row in df_imp_gs.iterrows():
                    try:
                        ot_raw = get_val(row, 'O T')
                        if not ot_raw:
                            continue
                        ot_val = f"OT-{ot_raw}"
                        # Solo importar OTs nuevas
                        if ot_val in ots_existentes:
                            continue

                        mod_val    = get_val(row, 'MODELO')
                        pat_val    = get_val(row, 'PATENTE')
                        col_val    = get_val(row, 'COLOR')
                        des_abr    = get_val(row, 'DES').upper()
                        pin_abr    = get_val(row, 'PIN').upper()
                        estado_abr = get_val(row, 'STADI', 'ESTADO').upper()
                        rep_val    = get_val(row, 'REPUESTOS')
                        # Daño: leer columna o usar mediano
                        dano_raw = get_val(row, 'DAÑO', 'DANO', 'TIPO DAÑO').lower()
                        dano_gs  = dano_raw if dano_raw in ('leve', 'mediano', 'grave') else 'mediano'

                        des_team = DES_MAP.get(des_abr, 'Rojas')
                        pin_team = PIN_MAP.get(pin_abr, None)  # None si no está asignado
                        estado   = ESTADO_MAP.get(estado_abr, 'Programado')
                        rep_str  = 'Completos ✅' if rep_val and rep_val not in ('nan','') else 'Pendiente ⏳'
                        # Si no hay equipo asignado, usar default para programar
                        des_team_prog = des_team if des_team else 'Rojas'
                        pin_team_prog = pin_team if pin_team else 'Percy PIN'

                        # Usar TALLER, si está vacío usar LLEGA, si no hay ninguna usar hoy
                        fecha_taller = None
                        for col_fecha in ['TALLER', 'LLEGA']:
                            try:
                                val = row.get(col_fecha)
                                if pd.notna(val) and str(val).strip() not in ('', 'nan', 'NaT'):
                                    fecha_taller = pd.to_datetime(val).date()
                                    break
                            except:
                                continue
                        if fecha_taller is None:
                            fecha_taller = date.today()

                        nombre_val = str(row.get('NOMBRE', '') or '').strip()
                        cia_val    = str(row.get('CIA', '') or '').strip()
                        liq_val    = str(row.get('LIQUIDADOR', '') or '').strip()
                        ase_val    = str(row.get('ASE', '') or '').strip()
                        loc_val    = str(row.get('LOC', '') or '').strip()
                        com_val    = str(row.get('COMENTARIO', '') or '').strip()
                        try:
                            llega_val = pd.to_datetime(row.get('LLEGA')).strftime('%Y-%m-%d')
                        except:
                            llega_val = ''

                        asigs = {}
                        auto_listo_idx = (fecha_taller - FECHA_BASE).days
                        es_listo = estado in ('Listo para Entrega',)
                        etapas_completadas_imp = set(ETAPAS_COMPLETADAS_POR_ESTADO.get(estado, []))
                        equipo_pin = pin_team
                        idx_f = auto_listo_idx

                        if es_listo:
                            for s in STAGES_ORDER:
                                if s in ('desabolladura','armado'): tc = des_team or 'Sin asignar'
                                elif s == 'pintura': tc = pin_team or 'Sin asignar'
                                elif s == 'pulido': tc = pul_de(pin_team) if pin_team else 'Sin asignar'
                                else: tc = 'Interno'
                                asigs[s] = {'team': tc,
                                    'start': fecha_taller.strftime('%Y-%m-%d'),
                                    'end': fecha_taller.strftime('%Y-%m-%d'),
                                    'eficiente': True, 'completado': True}
                        else:
                            for s in STAGES_ORDER:
                                g = STAGE_TO_GROUP[s]
                                dur = DURACIONES[dano_gs].get(g, 1)
                                if s in etapas_completadas_imp:
                                    if s in ('desabolladura','armado'): tc = des_team or 'Sin asignar'
                                    elif s == 'pintura': tc = pin_team or 'Sin asignar'
                                    elif s == 'pulido': tc = pul_de(pin_team) if pin_team else 'Sin asignar'
                                    else: tc = 'Interno'
                                    asigs[s] = {'team': tc,
                                        'start': fecha_taller.strftime('%Y-%m-%d'),
                                        'end': fecha_taller.strftime('%Y-%m-%d'),
                                        'eficiente': True, 'completado': True}
                                    idx_f = auto_listo_idx
                                    continue
                                if s in ('terminaciones', 'lavado'):
                                    idx_i, idx_f = buscar_bloque_libre(None, set(), auto_listo_idx, dur)
                                    asigs[s] = {'team': 'Interno',
                                        'start': (FECHA_BASE+timedelta(days=idx_i)).strftime('%Y-%m-%d'),
                                        'end': (FECHA_BASE+timedelta(days=idx_f)).strftime('%Y-%m-%d'),
                                        'eficiente': True}
                                else:
                                    if s == 'desabolladura': t_sel = des_team_prog
                                    elif s == 'pintura': t_sel = pin_team_prog
                                    elif s == 'armado': t_sel = des_team_prog
                                    elif s == 'pulido': t_sel = pul_de(pin_team_prog)
                                    else: t_sel = des_team_prog
                                    cal_a = temp_cals_gs.get(t_sel, {})
                                    idx_i, idx_f = buscar_bloque_libre(t_sel, cal_a, auto_listo_idx, dur)
                                    eff = verificar_eficiencia(t_sel, g, auto_listo_idx, dur, temp_cals_gs)
                                    if t_sel in TEAMS:
                                        for i in range(idx_i, idx_f+1):
                                            if es_dia_laboral(FECHA_BASE+timedelta(days=i)):
                                                temp_cals_gs[t_sel][i] = temp_cals_gs[t_sel].get(i,0)+1
                                    # Guardar equipo real (None si no asignado) para mostrar en tablero
                                    team_real = (des_team if s in ('desabolladura','armado')
                                                 else pin_team if s == 'pintura'
                                                 else pul_de(pin_team) if s == 'pulido' and pin_team
                                                 else t_sel)
                                    asigs[s] = {'team': team_real or 'Sin asignar',
                                        'start': (FECHA_BASE+timedelta(days=idx_i)).strftime('%Y-%m-%d'),
                                        'end': (FECHA_BASE+timedelta(days=idx_f)).strftime('%Y-%m-%d'),
                                        'eficiente': eff}
                                auto_listo_idx = idx_f + 1

                        fecha_entrega = (FECHA_BASE+timedelta(days=idx_f)).strftime('%Y-%m-%d')
                        db.upsert({
                            'ot': ot_val, 'modelo': mod_val, 'patente': pat_val,
                            'color': col_val, 'tipo_dano': dano_gs,
                            'asignaciones_json': json.dumps(asigs),
                            'fecha_entrega': fecha_entrega, 'estado': estado,
                            'repuestos': rep_str, 'nombre': nombre_val,
                            'cia': cia_val, 'liquidador': liq_val, 'ase': ase_val,
                            'loc': loc_val, 'comentario': com_val, 'fecha_llega': llega_val
                        })
                        importados_gs += 1
                    except Exception as e:
                        import traceback
                        errores_gs.append(f"OT-{get_val(row,'O T')}: {str(e)} | {traceback.format_exc().splitlines()[-1]}")

                st.session_state.vehicles = load_vehicles()
                st.session_state.calendarios = rebuild_calendars(st.session_state.vehicles)
                msg = f"✅ {importados_gs} vehículos nuevos importados"
                if eliminadas_gs > 0:
                    msg += f" | 🗑️ {eliminadas_gs} eliminados (ya no están en el Sheet)"
                st.success(msg)
                if errores_gs:
                    with st.expander(f"⚠️ {len(errores_gs)} filas con problemas — clic para ver detalle"):
                        for e in errores_gs:
                            st.caption(e)
                st.rerun()

        st.markdown("---")
    else:
        st.warning("⚠️ gspread no instalado. Ejecuta: `pip install gspread google-auth`")

    st.markdown("**📁 O importar desde Excel**")
    archivo = st.file_uploader("Sube tu Excel", type=["xlsx", "xls"])
    if archivo:
        try:
            df_imp = pd.read_excel(archivo)
            df_imp.columns = [str(c).strip() for c in df_imp.columns]

            # Preview
            st.write(f"**{len(df_imp)} filas encontradas.** Vista previa:")
            st.dataframe(df_imp.head(5), use_container_width=True, hide_index=True)

            dano_default = st.selectbox("Tipo de daño por defecto (si no hay columna DAÑO)", ["mediano", "leve", "grave"])

            if st.button("📥 Importar vehículos", type="primary"):
                importados = 0
                errores = []
                temp_cals = rebuild_calendars(st.session_state.vehicles)

                for _, row in df_imp.iterrows():
                    try:
                        ot_val   = f"OT-{str(row.get('O T', '')).strip()}"
                        mod_val  = str(row.get('MODELO', '')).strip()
                        pat_val  = str(row.get('PATENTE', '')).strip()
                        col_val  = str(row.get('COLOR', '')).strip()
                        des_abr  = str(row.get('DES', '')).strip().upper()
                        pin_abr  = str(row.get('PIN', '')).strip().upper()
                        estado_abr = str(row.get('ESTADO', '')).strip().upper()
                        rep_val  = str(row.get('REPUESTOS', '')).strip()
                        # Leer daño desde columna si existe, si no usar el default
                        dano_raw = str(row.get('DAÑO', row.get('DANO', row.get('TIPO DAÑO', '')))).strip().lower()
                        dano_default = dano_raw if dano_raw in ('leve', 'mediano', 'grave') else dano_default

                        des_team = DES_MAP.get(des_abr, 'Rojas')
                        pin_team = PIN_MAP.get(pin_abr, 'Percy')
                        estado   = ESTADO_MAP.get(estado_abr, 'Programado')
                        rep_str  = 'Completos ✅' if rep_val and rep_val != 'nan' else 'Pendiente ⏳'

                        # Calcular disponibilidad desde fecha TALLER
                        try:
                            fecha_taller = pd.to_datetime(row.get('TALLER')).date()
                        except:
                            fecha_taller = date.today()
                        disponibilidad_idx = (fecha_taller - FECHA_BASE).days

                        # Si ya está listo, marcar todas las etapas como completadas
                        asigs = {}
                        auto_listo_idx = disponibilidad_idx
                        es_listo = estado in ('Listo para Entrega', 'OK.')

                        if es_listo:
                            for s in STAGES_ORDER:
                                # Asignar equipo correcto segun etapa
                                if s in ('desabolladura', 'armado'):
                                    team_comp = des_team
                                elif s == 'pintura':
                                    team_comp = equipo_pin
                                elif s == 'pulido':
                                    team_comp = pul_de(equipo_pin)
                                else:
                                    team_comp = 'Interno'
                                asigs[s] = {
                                    'team': team_comp,
                                    'start': fecha_taller.strftime('%Y-%m-%d'),
                                    'end': fecha_taller.strftime('%Y-%m-%d'),
                                    'eficiente': True,
                                    'completado': True
                                }
                            idx_f = disponibilidad_idx
                        else:
                            equipo_des = des_team
                            equipo_pin = pin_team

                            # Determinar hasta qué etapa está completada según estado
                            # Secuencia: Desabolladura -> Pintura -> Armado -> Pulido -> Terminaciones -> Lavado
                            # El estado actual esta EN CURSO, solo las etapas ANTERIORES son completadas
                            ESTADO_ETAPA_COMPLETADA = {
                                'En Desabolladura': [],
                                'En Pintura':       ['desabolladura'],
                                'En Armado':        ['desabolladura', 'pintura'],
                                'En Pulido':        ['desabolladura', 'pintura', 'armado'],
                                'En Calidad':       ['desabolladura', 'pintura', 'armado', 'pulido'],
                                'Lavado':           ['desabolladura', 'pintura', 'armado', 'pulido', 'terminaciones'],
                                'Listo para Entrega': ['desabolladura', 'pintura', 'armado', 'pulido', 'terminaciones', 'lavado'],
                            }
                            etapas_completadas = set(ESTADO_ETAPA_COMPLETADA.get(estado, []))

                            for s in STAGES_ORDER:
                                g = STAGE_TO_GROUP[s]
                                dur = DURACIONES[dano_default].get(g, 1)

                                if s in etapas_completadas:
                                    if s in ('desabolladura', 'armado'):
                                        team_comp = des_team
                                    elif s == 'pintura':
                                        team_comp = equipo_pin
                                    elif s == 'pulido':
                                        team_comp = pul_de(equipo_pin)
                                    else:
                                        team_comp = 'Interno'
                                    asigs[s] = {
                                        'team': team_comp,
                                        'start': fecha_taller.strftime('%Y-%m-%d'),
                                        'end': fecha_taller.strftime('%Y-%m-%d'),
                                        'eficiente': True,
                                        'completado': True
                                    }
                                    disponibilidad_idx = max(disponibilidad_idx, (fecha_taller - FECHA_BASE).days)
                                    idx_f = disponibilidad_idx
                                    continue

                                if s in ('terminaciones', 'lavado'):
                                    idx_i, idx_f = buscar_bloque_libre(None, set(), disponibilidad_idx, dur)
                                    asigs[s] = {
                                        'team': 'Interno',
                                        'start': (FECHA_BASE + timedelta(days=idx_i)).strftime('%Y-%m-%d'),
                                        'end': (FECHA_BASE + timedelta(days=idx_f)).strftime('%Y-%m-%d'),
                                        'eficiente': True
                                    }
                                else:
                                    if s == 'armado':
                                        t_sel = equipo_des
                                    elif s == 'pulido':
                                        t_sel = equipo_pin
                                    elif s == 'desabolladura':
                                        t_sel = equipo_des
                                    elif s == 'pintura':
                                        t_sel = equipo_pin
                                    else:
                                        t_sel = equipo_des

                                    cal_actual = temp_cals.get(t_sel, set())
                                    idx_i, idx_f = buscar_bloque_libre(t_sel, cal_actual, disponibilidad_idx, dur)
                                    eff = verificar_eficiencia(t_sel, g, disponibilidad_idx, dur, temp_cals)

                                    if t_sel in TEAMS:
                                        for i in range(idx_i, idx_f + 1):
                                            if es_dia_laboral(FECHA_BASE + timedelta(days=i)):
                                                temp_cals[t_sel].add(i)

                                    asigs[s] = {
                                        'team': t_sel,
                                        'start': (FECHA_BASE + timedelta(days=idx_i)).strftime('%Y-%m-%d'),
                                        'end': (FECHA_BASE + timedelta(days=idx_f)).strftime('%Y-%m-%d'),
                                        'eficiente': eff
                                    }
                                disponibilidad_idx = idx_f + 1

                        fecha_entrega = (FECHA_BASE + timedelta(days=idx_f)).strftime('%Y-%m-%d')

                        nombre_val   = str(row.get('NOMBRE', '') or '').strip()
                        cia_val      = str(row.get('CIA', '') or '').strip()
                        liq_val      = str(row.get('LIQUIDADOR', '') or '').strip()
                        ase_val      = str(row.get('ASE', '') or '').strip()
                        loc_val      = str(row.get('LOC', '') or '').strip()
                        com_val      = str(row.get('COMENTARIO', '') or '').strip()
                        try:
                            llega_val = pd.to_datetime(row.get('LLEGA')).strftime('%Y-%m-%d')
                        except:
                            llega_val = ''

                        db.upsert({
                            'ot': ot_val, 'modelo': mod_val, 'patente': pat_val,
                            'color': col_val, 'tipo_dano': dano_default,
                            'asignaciones_json': json.dumps(asigs),
                            'fecha_entrega': fecha_entrega, 'estado': estado,
                            'repuestos': rep_str, 'nombre': nombre_val,
                            'cia': cia_val, 'liquidador': liq_val, 'ase': ase_val,
                            'loc': loc_val, 'comentario': com_val, 'fecha_llega': llega_val
                        })
                        importados += 1

                    except Exception as e:
                        errores.append(f"Fila {_ + 2}: {e}")

                st.session_state.vehicles = load_vehicles()
                st.session_state.calendarios = rebuild_calendars(st.session_state.vehicles)

                st.success(f"✅ {importados} vehículos importados correctamente")
                if errores:
                    st.warning("⚠️ Algunas filas tuvieron errores:")
                    for e in errores:
                        st.caption(e)
                st.rerun()

        except Exception as e:
            st.error(f"Error leyendo el archivo: {e}")

    st.markdown("---")
    if st.button("⚠️ REINICIAR TODO (elimina todos los vehículos)"):
        db.delete_all()
        st.session_state.vehicles = []
        st.session_state.calendarios = rebuild_calendars([])
        st.rerun()

# ── TAB 5: ALERTAS ──
with tab_alertas:
    st.subheader("🔔 Vencimientos")
    if alertas:
        for a in alertas:
            if "🚨" in a:
                st.error(a)
            else:
                st.warning(a)
    else:
        st.success("✅ Sin alertas de vencimiento.")

    st.markdown("---")
    st.subheader("⚠️ Asignaciones no óptimas")
    st.caption("Existe otro operario disponible antes que el asignado. Considera reasignar en Gestión.")
    ineficientes = []
    for v in st.session_state.vehicles:
        if v['Estado'] == "Listo para Entrega":
            continue
        for etapa, info in v['Asignaciones'].items():
            if not info.get('completado') and not info.get('eficiente', True):
                ineficientes.append({
                    'OT': v['OT'], 'Patente': v['Patente'], 'Modelo': v['Modelo'],
                    'Etapa': etapa.capitalize(),
                    'Equipo asignado': info.get('team', '—'),
                    'Inicio': info['start'].strftime('%d/%m/%Y') if info.get('start') else '—',
                })
    if ineficientes:
        st.dataframe(pd.DataFrame(ineficientes), use_container_width=True, hide_index=True)
    else:
        st.success("✅ Todas las asignaciones son óptimas.")

# ────────────────────────────────────────────────
# SECCIÓN INFERIOR COLAPSABLE: NUEVA OT + SEMÁFORO
# ────────────────────────────────────────────────
st.markdown("---")
with st.expander("➕ Nueva OT  |  🚥 Semáforo de Carga", expanded=False):
    col_form, col_sem = st.columns([3, 1])

    with col_form:
        st.subheader("⚙️ Programar Nueva OT")
        f1, f2, f3, f4, f5 = st.columns(5)
        ot_in   = f1.text_input("OT #", value=f"OT-{len(st.session_state.vehicles)+1:03d}", key="f_ot")
        pat_in  = f2.text_input("Patente", key="f_pat")
        mod_in  = f3.text_input("Modelo", key="f_mod")
        col_in  = f4.text_input("Color", key="f_col")
        dano_in = f5.selectbox("Daño", ["leve", "mediano", "grave"], key="f_dano")

        with st.expander("➕ Datos adicionales del cliente"):
            g1, g2, g3, g4 = st.columns(4)
            nom_in   = g1.text_input("Nombre cliente", key="f_nom")
            cia_in   = g2.text_input("CIA", key="f_cia")
            liq_in   = g3.text_input("Liquidador", key="f_liq")
            ase_in   = g4.text_input("ASE", key="f_ase")
            g5, g6, g7 = st.columns(3)
            loc_in   = g5.text_input("LOC", key="f_loc")
            llega_in = g6.date_input("Fecha llega", value=date.today(), key="f_llega")
            com_in   = g7.text_area("Comentario", height=60, key="f_com")

        m1, m2 = st.columns([1, 2])
        modo_auto = m1.toggle("🤖 Asignación Automática", value=True, key="f_auto")
        cho_teams = {}
        if not modo_auto:
            ma1, ma2 = st.columns(2)
            cho_teams['desabolladura'] = ma1.selectbox("Desabolladura / Armado", ROTATION_ORDER['desabolladura'], key="f_des")
            cho_teams['pintura']       = ma2.selectbox("Pintura (Pulido automático)", ROTATION_ORDER['pintura'], key="f_pin")

        if st.button("🚀 Programar Vehículo", type="primary", key="btn_programar"):
            asigs = {}
            temp_cals = rebuild_calendars(st.session_state.vehicles)
            hoy_idx = (date.today() - FECHA_BASE).days
            auto_listo_idx = hoy_idx
            equipo_desabolladura = None
            equipo_pintura = None

            for s in STAGES_ORDER:
                g = STAGE_TO_GROUP[s]
                dur = DURACIONES[dano_in].get(g, 1)
                if s in ('terminaciones', 'lavado'):
                    idx_i, idx_f = buscar_bloque_libre(None, set(), auto_listo_idx, dur)
                    asigs[s] = {'team': 'Interno',
                                'start': (FECHA_BASE + timedelta(days=idx_i)).strftime('%Y-%m-%d'),
                                'end':   (FECHA_BASE + timedelta(days=idx_f)).strftime('%Y-%m-%d'),
                                'eficiente': True}
                    auto_listo_idx = idx_f + 1
                    continue
                if s == 'armado':
                    t_sel = equipo_desabolladura
                elif s == 'pulido':
                    t_sel = pul_de(equipo_pintura)
                elif modo_auto:
                    t_sel = elegir_equipo_optimo(g, auto_listo_idx, dur, temp_cals)
                else:
                    t_sel = cho_teams.get(g, 'Rojas')
                if s == 'desabolladura': equipo_desabolladura = t_sel
                elif s == 'pintura':     equipo_pintura = t_sel
                cal_actual = temp_cals.get(t_sel, set())
                idx_i, idx_f = buscar_bloque_libre(t_sel, cal_actual, auto_listo_idx, dur)
                eff = verificar_eficiencia(t_sel, g, auto_listo_idx, dur, temp_cals)
                if t_sel in TEAMS:
                    for i in range(idx_i, idx_f + 1):
                        if es_dia_laboral(FECHA_BASE + timedelta(days=i)):
                            temp_cals[t_sel].add(i)
                asigs[s] = {'team': t_sel,
                            'start': (FECHA_BASE + timedelta(days=idx_i)).strftime('%Y-%m-%d'),
                            'end':   (FECHA_BASE + timedelta(days=idx_f)).strftime('%Y-%m-%d'),
                            'eficiente': eff}
                auto_listo_idx = idx_f + 1

            fecha_entrega_str = (FECHA_BASE + timedelta(days=idx_f)).strftime('%Y-%m-%d')
            db.upsert({
                'ot': ot_in, 'modelo': mod_in, 'patente': pat_in,
                'color': col_in, 'tipo_dano': dano_in,
                'asignaciones_json': json.dumps(asigs),
                'fecha_entrega': fecha_entrega_str, 'estado': 'Programado',
                'repuestos': 'Pendiente ⏳', 'nombre': nom_in,
                'cia': cia_in, 'liquidador': liq_in, 'ase': ase_in,
                'loc': loc_in, 'comentario': com_in,
                'fecha_llega': llega_in.strftime('%Y-%m-%d')
            })
            st.session_state.vehicles = load_vehicles()
            st.session_state.calendarios = rebuild_calendars(st.session_state.vehicles)
            st.success(f"✅ {ot_in} programado — Entrega estimada: {fecha_entrega_str}")
            st.rerun()

    with col_sem:
        st.subheader("🚥 Semáforo (14 días)")
        for t in TEAMS:
            cal = st.session_state.calendarios.get(t, {})
            cap = CAPACIDAD.get(t, 1)
            hoy_idx = (date.today() - FECHA_BASE).days
            dias_hab_14 = sum(1 for i in range(hoy_idx, hoy_idx + 14)
                              if es_dia_laboral(FECHA_BASE + timedelta(days=i)))
            ocupacion = sum(cal.get(i, 0) for i in range(hoy_idx, hoy_idx + 14))
            max_posible = cap * dias_hab_14
            pct = ocupacion / max_posible if max_posible > 0 else 0
            if pct == 0:    sem = "⚪ LIBRE"
            elif pct < 0.4: sem = "🟢 LIBRE"
            elif pct < 0.8: sem = "🟡 NORMAL"
            else:           sem = "🔴 SATURADO"
            st.write(f"**{t}**: {sem} ({round(pct*100)}%)")
