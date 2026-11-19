from datetime import date, datetime
import uuid
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="CobroPro | Gestión de Préstamos", layout="wide"
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

COLS_CAJA = [
    "fecha",
    "monto_inicial",
    "cobrador",
    "estado",
    "timestamp",
]
COLS_PRESTAMOS = [
    "id_prestamo",
    "cliente",
    "telefono",
    "monto_prestado",
    "interes_pct",
    "total_a_pagar",
    "cuota_valor",
    "frecuencia",
    "saldo_pendiente",
    "fecha_inicio",
    "estado",
]
COLS_PAGOS = [
    "id_pago",
    "id_prestamo",
    "cliente",
    "monto_pagado",
    "metodo_pago",
    "saldo_restante",
    "fecha_hora",
    "cobrador",
]
COLS_MOVS = [
    "id_mov",
    "tipo",
    "referencia",
    "monto",
    "detalle",
    "fecha_hora",
]


@st.cache_resource
def get_spreadsheet():
  creds = Credentials.from_service_account_info(
      st.secrets["gcp_service_account"], scopes=SCOPES
  )
  client = gspread.authorize(creds)
  return client.open_by_key(st.secrets["SHEET_ID"])


def ensure_ws(sh, title, headers):
  try:
    ws = sh.worksheet(title)
  except gspread.WorksheetNotFound:
    ws = sh.add_worksheet(title=title, rows=2000, cols=len(headers) + 2)
    ws.append_row(headers)
  values = ws.get_all_values()
  if not values:
    ws.append_row(headers)
  return ws


def read_df(ws, headers):
  records = ws.get_all_records()
  if not records:
    return pd.DataFrame(columns=headers)
  df = pd.DataFrame(records)
  for col in headers:
    if col not in df.columns:
      df[col] = None
  return df[headers]


def append_movimiento(ws_mov, tipo, referencia, monto, detalle):
  ws_mov.append_row([
      f"MOV-{uuid.uuid4().hex[:8].upper()}",
      tipo,
      referencia,
      float(monto),
      detalle,
      datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
  ])


# --- Conexión y preparación de hojas ---
sh = get_spreadsheet()
ws_caja = ensure_ws(sh, "caja_inicial", COLS_CAJA)
ws_prestamos = ensure_ws(sh, "prestamos", COLS_PRESTAMOS)
ws_pagos = ensure_ws(sh, "historial_pagos", COLS_PAGOS)
ws_movs = ensure_ws(sh, "historial_movimientos", COLS_MOVS)

df_caja = read_df(ws_caja, COLS_CAJA)
df_prestamos = read_df(ws_prestamos, COLS_PRESTAMOS)
df_pagos = read_df(ws_pagos, COLS_PAGOS)
df_movs = read_df(ws_movs, COLS_MOVS)

hoy_str = date.today().strftime("%Y-%m-%d")

# --- Sidebar: Cuadre de Caja Inicial Único ---
with st.sidebar:
  st.header("Cuadre de Caja Inicial")
  cobrador_activo = st.text_input("Nombre del cobrador", value="Cobrador 01")

  caja_hoy = (
      df_caja[df_caja["fecha"].astype(str) == hoy_str]
      if not df_caja.empty
      else pd.DataFrame()
  )

  if not caja_hoy.empty:
    monto_caja_inicial = float(caja_hoy.iloc[0]["monto_inicial"])
    st.success(f"Caja abierta hoy ({hoy_str}): **${monto_caja_inicial:,.2f}**")
    st.caption(
        "Solo se permite un cuadre inicial por fecha para mantener la"
        " integridad contable."
    )
  else:
    monto_caja_inicial = 0.0
    st.warning("Aún no has registrado el cuadre de caja inicial de hoy.")
    with st.form("form_caja_inicial"):
      monto_input = st.number_input(
          "Monto inicial en caja ($)", min_value=0.0, step=10.0, format="%.2f"
      )
      submitted_caja = st.form_submit_button("Registrar Cuadre Inicial")
      if submitted_caja:
        ws_caja.append_row([
            hoy_str,
            float(monto_input),
            cobrador_activo.strip() or "Cobrador",
            "ABIERTA",
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ])
        append_movimiento(
            ws_movs,
            "APERTURA_CAJA",
            hoy_str,
            monto_input,
            f"Cuadre inicial por {cobrador_activo}",
        )
        st.rerun()

# --- Cálculos del día para el cuadre en tiempo real ---
pagos_hoy = (
    df_pagos[
        df_pagos["fecha_hora"].astype(str).str.startswith(hoy_str)
    ].copy()
    if not df_pagos.empty
    else pd.DataFrame(columns=COLS_PAGOS)
)
total_cobrado_hoy = (
    pd.to_numeric(pagos_hoy["monto_pagado"], errors="coerce").fillna(0).sum()
    if not pagos_hoy.empty
    else 0.0
)

prestamos_hoy = (
    df_prestamos[
        df_prestamos["fecha_inicio"].astype(str) == hoy_str
    ].copy()
    if not df_prestamos.empty
    else pd.DataFrame(columns=COLS_PRESTAMOS)
)
total_prestado_hoy = (
    pd.to_numeric(prestamos_hoy["monto_prestado"], errors="coerce")
    .fillna(0)
    .sum()
    if not prestamos_hoy.empty
    else 0.0
)

efectivo_esperado = monto_caja_inicial + total_cobrado_hoy - total_prestado_hoy
cartera_activa = (
    pd.to_numeric(df_prestamos["saldo_pendiente"], errors="coerce")
    .fillna(0)
    .sum()
    if not df_prestamos.empty
    else 0.0
)

# --- Encabezado y KPIs ---
st.title("Sistema de Cobro de Préstamos Personales")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Caja Inicial (Hoy)", f"${monto_caja_inicial:,.2f}")
c2.metric("Cobrado Hoy", f"${total_cobrado_hoy:,.2f}")
c3.metric("Prestado Hoy", f"${total_prestado_hoy:,.2f}")
c4.metric("Efectivo Esperado en Caja", f"${efectivo_esperado:,.2f}")

st.caption(
    f"Cartera activa total pendiente por cobrar: **${cartera_activa:,.2f}**"
)

if caja_hoy.empty:
  st.info(
      "Registra primero el **Cuadre de Caja Inicial** en la barra lateral para"
      " operar la jornada."
  )
  st.stop()

# --- Pestañas operativas del cobrador ---
tab_cobro, tab_nuevo, tab_cartera, tab_historial = st.tabs([
    "Registrar Cobro / Ruta",
    "Nuevo Préstamo",
    "Cartera de Clientes",
    "Historial Inmutable (Sheets)",
])

with tab_cobro:
  st.subheader("Registrar Abono o Pago de Cliente")
  activos = df_prestamos[
      pd.to_numeric(df_prestamos["saldo_pendiente"], errors="coerce").fillna(0)
      > 0
  ].copy()

  if activos.empty:
    st.write("No hay préstamos con saldo pendiente en este momento.")
  else:
    activos["etiqueta"] = activos.apply(
        lambda r: (
            f"{r['id_prestamo']} | {r['cliente']} | Saldo:"
            f" ${float(r['saldo_pendiente']):,.2f} | Cuota:"
            f" ${float(r['cuota_valor']):,.2f}"
        ),
        axis=1,
    )
    seleccion = st.selectbox("Selecciona el préstamo", activos["etiqueta"].tolist())
    fila_p = activos[activos["etiqueta"] == seleccion].iloc[0]
    saldo_actual = float(fila_p["saldo_pendiente"])
    cuota_sugerida = min(float(fila_p["cuota_valor"]), saldo_actual)

    with st.form("form_pago"):
      monto_pago = st.number_input(
          "Monto a cobrar ($)",
          min_value=0.01,
          max_value=saldo_actual,
          value=max(0.01, cuota_sugerida),
          step=1.0,
          format="%.2f",
      )
      metodo = st.selectbox(
          "Método de pago", ["Efectivo", "Pago Móvil / Transferencia", "Otro"]
      )
      confirmar_pago = st.form_submit_button("Guardar Pago en Historial")

      if confirmar_pago:
        nuevo_saldo = round(saldo_actual - float(monto_pago), 2)
        nuevo_estado = "PAGADO" if nuevo_saldo <= 0 else "ACTIVO"
        id_pago = f"PAG-{uuid.uuid4().hex[:8].upper()}"
        ahora_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 1. Append inmutable al historial de pagos
        ws_pagos.append_row([
            id_pago,
            fila_p["id_prestamo"],
            fila_p["cliente"],
            float(monto_pago),
            metodo,
            nuevo_saldo,
            ahora_str,
            cobrador_activo,
        ])

        # 2. Actualizar únicamente el saldo del préstamo en 'prestamos' (nunca se borra historial)
        cell = ws_prestamos.find(fila_p["id_prestamo"], in_column=1)
        if cell:
          ws_prestamos.update_cell(cell.row, 9, nuevo_saldo)  # saldo_pendiente
          ws_prestamos.update_cell(cell.row, 11, nuevo_estado)  # estado

        append_movimiento(
            ws_movs,
            "COBRO",
            id_pago,
            monto_pago,
            f"Abono a {fila_p['id_prestamo']} ({fila_p['cliente']}). Nuevo"
            f" saldo: ${nuevo_saldo:,.2f}",
        )
        st.success(
            f"Pago {id_pago} registrado. Saldo restante: ${nuevo_saldo:,.2f}"
        )
        st.rerun()

with tab_nuevo:
  st.subheader("Otorgar Nuevo Préstamo")
  with st.form("form_nuevo_prestamo"):
    col_a, col_b = st.columns(2)
    cliente = col_a.text_input("Nombre completo del cliente")
    telefono = col_b.text_input("Teléfono / WhatsApp")

    col_c, col_d, col_e = st.columns(3)
    monto_prestado = col_c.number_input(
        "Capital prestado ($)", min_value=1.0, value=100.0, step=10.0
    )
    interes_pct = col_d.number_input(
        "Interés total (%)", min_value=0.0, value=20.0, step=1.0
    )
    num_cuotas = col_e.number_input(
        "Número de cuotas", min_value=1, value=20, step=1
    )

    frecuencia = st.selectbox(
        "Frecuencia de cobro", ["Diario", "Semanal", "Quincenal", "Mensual"]
    )
    crear_p = st.form_submit_button("Registrar Préstamo")

    if crear_p:
      if not cliente.strip():
        st.error("Debes ingresar el nombre del cliente.")
      else:
        total_pagar = round(
            float(monto_prestado) * (1 + float(interes_pct) / 100.0), 2
        )
        cuota_val = round(total_pagar / int(num_cuotas), 2)
        id_prestamo = f"PRE-{uuid.uuid4().hex[:7].upper()}"

        ws_prestamos.append_row([
            id_prestamo,
            cliente.strip(),
            telefono.strip(),
            float(monto_prestado),
            float(interes_pct),
            total_pagar,
            cuota_val,
            frecuencia,
            total_pagar,
            hoy_str,
            "ACTIVO",
        ])
        append_movimiento(
            ws_movs,
            "NUEVO_PRESTAMO",
            id_prestamo,
            monto_prestado,
            f"Cliente: {cliente.strip()} | Total a pagar: ${total_pagar:,.2f}",
        )
        st.success(
            f"Préstamo {id_prestamo} creado. Total a pagar: ${total_pagar:,.2f}"
            f" ({num_cuotas} cuotas de ${cuota_val:,.2f})"
        )
        st.rerun()

with tab_cartera:
  st.subheader("Cartera Completa de Préstamos")
  filtro_estado = st.selectbox("Filtrar por estado", ["TODOS", "ACTIVO", "PAGADO"])
  vista_p = df_prestamos.copy()
  if filtro_estado != "TODOS" and not vista_p.empty:
    vista_p = vista_p[vista_p["estado"] == filtro_estado]
  st.dataframe(vista_p, use_container_width=True, hide_index=True)

with tab_historial:
  st.subheader("Historial de Pagos (Solo Lectura / Nunca se Borra)")
  st.dataframe(df_pagos, use_container_width=True, hide_index=True)

  st.subheader("Bitácora de Movimientos del Sistema")
  st.dataframe(df_movs, use_container_width=True, hide_index=True)
