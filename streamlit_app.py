from datetime import date, datetime
import re
import uuid
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="CobroPro | Gestión de Préstamos",
    layout="wide",
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
COLS_CLIENTES = [
    "id_cliente",
    "nombre",
    "telefono",
    "direccion",
    "referencia",
    "fecha_registro",
    "estado",
]
COLS_PRESTAMOS = [
    "id_prestamo",
    "id_cliente",
    "cliente",
    "telefono",
    "monto_prestado",
    "interes_pct",
    "total_a_pagar",
    "cuota_valor",
    "frecuencia",
    "saldo_pendiente",
    "fecha_inicio",
    "afecta_caja",
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
  if "gcp_service_account" not in st.secrets or "SHEET_ID" not in st.secrets:
    st.error(
        "Falta configurar `gcp_service_account` o `SHEET_ID` en Streamlit"
        " Secrets."
    )
    st.stop()

  info = dict(st.secrets["gcp_service_account"])
  if "private_key" in info and isinstance(info["private_key"], str):
    info["private_key"] = info["private_key"].replace("\\n", "\n").strip()

  raw_sheet_id = str(st.secrets["SHEET_ID"]).strip()
  match = re.search(r"/d/([a-zA-Z0-9-_]+)", raw_sheet_id)
  clean_sheet_id = match.group(1) if match else raw_sheet_id

  try:
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(clean_sheet_id)
  except Exception as exc:
    st.error(
        "Error al autenticar o abrir Google Sheets. Revisa la `private_key`,"
        f" el `SHEET_ID` y los permisos de la hoja: {exc}"
    )
    st.stop()


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
  rows = ws.get_all_values()
  if not rows or len(rows) <= 1:
    return pd.DataFrame(columns=headers)

  raw_headers = [str(h).strip() for h in rows[0]]
  data_rows = rows[1:]

  if raw_headers == headers:
    cleaned = [
        r[: len(headers)] + [None] * max(0, len(headers) - len(r))
        for r in data_rows
        if any(str(c).strip() != "" for c in r)
    ]
    return pd.DataFrame(cleaned, columns=headers)

  safe_cols = [
      name if name else f"col_{idx}" for idx, name in enumerate(raw_headers)
  ]
  normalized_rows = []
  for r in data_rows:
    if not any(str(c).strip() != "" for c in r):
      continue
    padded = r[: len(safe_cols)] + [None] * max(0, len(safe_cols) - len(r))
    normalized_rows.append(dict(zip(safe_cols, padded)))

  if not normalized_rows:
    return pd.DataFrame(columns=headers)

  df = pd.DataFrame(normalized_rows)
  for col in headers:
    if col not in df.columns:
      df[col] = None
  return df[headers]


def parse_num(series):
  return pd.to_numeric(
      series.astype(str)
      .str.replace("$", "", regex=False)
      .str.replace(",", "", regex=False)
      .str.strip(),
      errors="coerce",
  ).fillna(0.0)


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
ws_clientes = ensure_ws(sh, "clientes", COLS_CLIENTES)
ws_prestamos = ensure_ws(sh, "prestamos", COLS_PRESTAMOS)
ws_pagos = ensure_ws(sh, "historial_pagos", COLS_PAGOS)
ws_movs = ensure_ws(sh, "historial_movimientos", COLS_MOVS)

df_caja = read_df(ws_caja, COLS_CAJA)
df_clientes = read_df(ws_clientes, COLS_CLIENTES)
df_prestamos = read_df(ws_prestamos, COLS_PRESTAMOS)
df_pagos = read_df(ws_pagos, COLS_PAGOS)
df_movs = read_df(ws_movs, COLS_MOVS)

hoy_str = date.today().strftime("%Y-%m-%d")

# --- Sidebar: Cuadre de Caja Inicial Único ---
with st.sidebar:
  st.header("Cuadre de Caja Inicial")
  cobrador_activo = st.text_input(
      "Nombre del cobrador",
      value="Cobrador 01",
      key="input_cobrador_activo",
  )

  caja_hoy = (
      df_caja[df_caja["fecha"].astype(str) == hoy_str]
      if not df_caja.empty
      else pd.DataFrame()
  )

  if not caja_hoy.empty:
    monto_caja_inicial = float(parse_num(caja_hoy["monto_inicial"]).iloc[0])
    st.success(f"Caja abierta hoy ({hoy_str}): **${monto_caja_inicial:,.2f}**")
    st.caption(
        "El cuadre inicial es único por día y no cambia al registrar clientes"
        " ni préstamos."
    )
  else:
    monto_caja_inicial = 0.0
    st.warning("Aún no has registrado el cuadre de caja inicial de hoy.")
    with st.form("form_caja_inicial", clear_on_submit=True):
      monto_input = st.number_input(
          "Monto inicial en caja ($)",
          min_value=0.0,
          step=10.0,
          format="%.2f",
          key="input_monto_caja_inicial",
      )
      submitted_caja = st.form_submit_button("Registrar Cuadre Inicial")
      if submitted_caja:
        nombre_cob = cobrador_activo.strip() or "Cobrador"
        ws_caja.append_row([
            hoy_str,
            float(monto_input),
            nombre_cob,
            "ABIERTA",
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ])
        append_movimiento(
            ws_movs,
            "APERTURA_CAJA",
            hoy_str,
            monto_input,
            f"Cuadre inicial por {nombre_cob}",
        )
        st.rerun()

# --- Cálculos del día ---
pagos_hoy = (
    df_pagos[df_pagos["fecha_hora"].astype(str).str.startswith(hoy_str)].copy()
    if not df_pagos.empty
    else pd.DataFrame(columns=COLS_PAGOS)
)
total_cobrado_hoy = (
    parse_num(pagos_hoy["monto_pagado"]).sum() if not pagos_hoy.empty else 0.0
)

prestamos_hoy = (
    df_prestamos[df_prestamos["fecha_inicio"].astype(str) == hoy_str].copy()
    if not df_prestamos.empty
    else pd.DataFrame(columns=COLS_PRESTAMOS)
)

if not prestamos_hoy.empty:
  prestamos_hoy["monto_num"] = parse_num(prestamos_hoy["monto_prestado"])
  afecta_mask = prestamos_hoy["afecta_caja"].astype(str).str.strip().str.upper() == "SI"
  total_prestado_caja_hoy = prestamos_hoy.loc[afecta_mask, "monto_num"].sum()
  total_prestado_fuera_caja_hoy = prestamos_hoy.loc[~afecta_mask, "monto_num"].sum()
else:
  total_prestado_caja_hoy = 0.0
  total_prestado_fuera_caja_hoy = 0.0

efectivo_esperado = (
    monto_caja_inicial + total_cobrado_hoy - total_prestado_caja_hoy
)
cartera_activa = (
    parse_num(df_prestamos["saldo_pendiente"]).sum()
    if not df_prestamos.empty
    else 0.0
)

# --- Encabezado y KPIs ---
st.title("Sistema de Cobro de Préstamos Personales")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Caja Inicial (Hoy)", f"${monto_caja_inicial:,.2f}")
c2.metric("Cobrado Hoy", f"${total_cobrado_hoy:,.2f}")
c3.metric("Prestado Hoy (En Caja)", f"${total_prestado_caja_hoy:,.2f}")
c4.metric("Efectivo Esperado en Caja", f"${efectivo_esperado:,.2f}")

st.caption(
    f"Prestado hoy sin mover caja: **${total_prestado_fuera_caja_hoy:,.2f}**"
    f" | Clientes registrados: **{len(df_clientes)}** | Cartera activa"
    f" pendiente: **${cartera_activa:,.2f}**"
)

if caja_hoy.empty:
  st.info(
      "Registra primero el **Cuadre de Caja Inicial** en la barra lateral para"
      " operar la jornada."
  )
  st.stop()

# --- Pestañas operativas ---
tab_cobro, tab_clientes, tab_nuevo_p, tab_cartera, tab_historial = st.tabs([
    "Registrar Cobro / Ruta",
    "Registrar Clientes",
    "Nuevo Préstamo",
    "Cartera de Clientes",
    "Historial Inmutable (Sheets)",
])

with tab_cobro:
  st.subheader("Registrar Abono o Pago de Cliente")
  contenedor_cobro = st.container()

  with contenedor_cobro:
    if df_prestamos.empty:
      st.info("No hay préstamos registrados todavía.")
    else:
      df_p_calc = df_prestamos.copy()
      df_p_calc["saldo_num"] = parse_num(df_p_calc["saldo_pendiente"])
      df_p_calc["cuota_num"] = parse_num(df_p_calc["cuota_valor"])
      df_p_calc["estado_limpio"] = (
          df_p_calc["estado"].astype(str).str.strip().str.upper()
      )

      activos = df_p_calc[
          (df_p_calc["saldo_num"] > 0) & (df_p_calc["estado_limpio"] != "PAGADO")
      ].copy()

      if activos.empty:
        st.info("No hay préstamos con saldo pendiente por cobrar.")
      else:
        activos["etiqueta"] = activos.apply(
            lambda r: (
                f"{r['id_prestamo']} | {r['cliente']} | Saldo por cobrar:"
                f" ${r['saldo_num']:,.2f} | Cuota: ${r['cuota_num']:,.2f} |"
                f" Caja: {r['afecta_caja'] or 'SI'}"
            ),
            axis=1,
        )
        seleccion = st.selectbox(
            "Selecciona el préstamo",
            activos["etiqueta"].tolist(),
            key="select_prestamo_cobro",
        )
        fila_p = activos[activos["etiqueta"] == seleccion].iloc[0]
        saldo_actual = float(fila_p["saldo_num"])
        cuota_sugerida = min(float(fila_p["cuota_num"]), saldo_actual)

        with st.form("form_pago", clear_on_submit=True):
          monto_pago = st.number_input(
              "Monto a cobrar ($)",
              min_value=0.01,
              max_value=saldo_actual,
              value=max(0.01, cuota_sugerida),
              step=1.0,
              format="%.2f",
              key="input_monto_pago",
          )
          metodo = st.selectbox(
              "Método de pago",
              ["Efectivo", "Pago Móvil / Transferencia", "Otro"],
              key="select_metodo_pago",
          )
          confirmar_pago = st.form_submit_button("Guardar Pago en Historial")

          if confirmar_pago:
            nuevo_saldo = round(saldo_actual - float(monto_pago), 2)
            nuevo_estado = "PAGADO" if nuevo_saldo <= 0 else "ACTIVO"
            id_pago = f"PAG-{uuid.uuid4().hex[:8].upper()}"
            ahora_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            nombre_cob = cobrador_activo.strip() or "Cobrador"

            ws_pagos.append_row([
                id_pago,
                fila_p["id_prestamo"],
                fila_p["cliente"],
                float(monto_pago),
                metodo,
                nuevo_saldo,
                ahora_str,
                nombre_cob,
            ])

            cell = ws_prestamos.find(fila_p["id_prestamo"], in_column=1)
            if cell:
              ws_prestamos.update_cell(cell.row, 10, nuevo_saldo)
              ws_prestamos.update_cell(cell.row, 13, nuevo_estado)

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

with tab_clientes:
  st.subheader("Registrar Nuevo Cliente (No altera la Caja Inicial)")
  with st.form("form_nuevo_cliente", clear_on_submit=True):
    c_nom, c_tel = st.columns(2)
    cli_nombre = c_nom.text_input("Nombre completo *", key="input_cli_nombre")
    cli_telefono = c_tel.text_input(
        "Teléfono / WhatsApp", key="input_cli_telefono"
    )

    c_dir, c_ref = st.columns(2)
    cli_direccion = c_dir.text_input(
        "Dirección / Sector de ruta", key="input_cli_direccion"
    )
    cli_referencia = c_ref.text_input(
        "Referencia / Nota", key="input_cli_referencia"
    )

    guardar_cliente = st.form_submit_button("Guardar Cliente")
    if guardar_cliente:
      if not cli_nombre.strip():
        st.error("El nombre del cliente es obligatorio.")
      else:
        id_cliente = f"CLI-{uuid.uuid4().hex[:7].upper()}"
        ws_clientes.append_row([
            id_cliente,
            cli_nombre.strip(),
            cli_telefono.strip(),
            cli_direccion.strip(),
            cli_referencia.strip(),
            hoy_str,
            "ACTIVO",
        ])
        append_movimiento(
            ws_movs,
            "NUEVO_CLIENTE",
            id_cliente,
            0.0,
            f"Cliente registrado: {cli_nombre.strip()} ({cli_telefono.strip()})",
        )
        st.success(
            f"Cliente **{cli_nombre.strip()}** ({id_cliente}) guardado sin"
            " modificar tu caja inicial."
        )
        st.rerun()

  st.markdown("### Directorio de Clientes")
  st.dataframe(df_clientes, use_container_width=True, hide_index=True)

with tab_nuevo_p:
  st.subheader("Otorgar Nuevo Préstamo a un Cliente")
  if df_clientes.empty:
    st.warning(
        "Primero registra al menos un cliente en la pestaña **Registrar"
        " Clientes**."
    )
  else:
    df_clientes_view = df_clientes.copy()
    df_clientes_view["selector"] = df_clientes_view.apply(
        lambda r: (
            f"{r['id_cliente']} | {r['nombre']} | Tel:"
            f" {r['telefono'] or 'N/A'}"
        ),
        axis=1,
    )
    cli_sel_str = st.selectbox(
        "Selecciona el cliente",
        df_clientes_view["selector"].tolist(),
        key="select_cliente_prestamo",
    )
    cli_fila = df_clientes_view[
        df_clientes_view["selector"] == cli_sel_str
    ].iloc[0]

    with st.form("form_nuevo_prestamo", clear_on_submit=True):
      col_c, col_d, col_e = st.columns(3)
      monto_prestado = col_c.number_input(
          "Capital prestado ($)",
          min_value=1.0,
          value=100.0,
          step=10.0,
          key="input_monto_prestado",
      )
      interes_pct = col_d.number_input(
          "Interés total (%)",
          min_value=0.0,
          value=20.0,
          step=1.0,
          key="input_interes_pct",
      )
      num_cuotas = col_e.number_input(
          "Número de cuotas",
          min_value=1,
          value=20,
          step=1,
          key="input_num_cuotas",
      )

      frecuencia = st.selectbox(
          "Frecuencia de cobro",
          ["Diario", "Semanal", "Quincenal", "Mensual"],
          key="select_frecuencia_cobro",
      )

      descontar_de_caja = st.toggle(
          "Descontar este préstamo de la caja del día (SI / NO)",
          value=True,
          key="toggle_descontar_caja",
      )

      crear_p = st.form_submit_button("Registrar Préstamo")

      if crear_p:
        total_pagar = round(
            float(monto_prestado) * (1 + float(interes_pct) / 100.0), 2
        )
        cuota_val = round(total_pagar / int(num_cuotas), 2)
        id_prestamo = f"PRE-{uuid.uuid4().hex[:7].upper()}"
        flag_caja = "SI" if descontar_de_caja else "NO"

        ws_prestamos.append_row([
            id_prestamo,
            cli_fila["id_cliente"],
            cli_fila["nombre"],
            cli_fila["telefono"],
            float(monto_prestado),
            float(interes_pct),
            total_pagar,
            cuota_val,
            frecuencia,
            total_pagar,
            hoy_str,
            flag_caja,
            "ACTIVO",
        ])
        append_movimiento(
            ws_movs,
            "NUEVO_PRESTAMO",
            id_prestamo,
            monto_prestado,
            f"Cliente: {cli_fila['nombre']} ({cli_fila['id_cliente']}) |"
            f" Total: ${total_pagar:,.2f} | Afecta caja: {flag_caja}",
        )
        st.success(
            f"Préstamo {id_prestamo} creado para **{cli_fila['nombre']}**."
            f" Saldo disponible por cobrar: **${total_pagar:,.2f}** | Afecta"
            f" caja: **{flag_caja}**"
        )
        st.rerun()

with tab_cartera:
  st.subheader("Cartera Completa de Préstamos")
  filtro_estado = st.selectbox(
      "Filtrar por estado",
      ["TODOS", "ACTIVO", "PAGADO"],
      key="select_filtro_estado",
  )
  vista_p = df_prestamos.copy()
  if filtro_estado != "TODOS" and not vista_p.empty:
    vista_p = vista_p[vista_p["estado"] == filtro_estado]
  st.dataframe(vista_p, use_container_width=True, hide_index=True)

with tab_historial:
  st.subheader("Historial de Pagos (Solo Lectura / Nunca se Borra)")
  st.dataframe(df_pagos, use_container_width=True, hide_index=True)

  st.subheader("Bitácora de Movimientos del Sistema")
  st.dataframe(df_movs, use_container_width=True, hide_index=True)
