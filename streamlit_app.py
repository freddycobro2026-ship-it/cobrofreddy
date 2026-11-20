from datetime import date, datetime
import re
import time
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
COLS_GASTOS = [
    "id_gasto",
    "categoria",
    "descripcion",
    "monto",
    "metodo_pago",
    "afecta_caja",
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


def ensure_ws(sh, title, headers, retries=3, delay=1.0):
  for attempt in range(retries):
    try:
      try:
        ws = sh.worksheet(title)
      except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=2000, cols=len(headers) + 4)
        ws.append_row(headers)
      values = ws.get_all_values()
      if not values:
        ws.append_row(headers)
      else:
        existing_headers = [str(h).strip() for h in values[0]]
        missing = [h for h in headers if h not in existing_headers]
        if missing:
          new_headers = existing_headers + missing
          ws.update(range_name="A1", values=[new_headers])
      return ws
    except Exception as exc:
      if attempt < retries - 1:
        time.sleep(delay * (2**attempt))
      else:
        st.error(
            f"Error preparando la pestaña '{title}' en Google Sheets: {exc}"
        )
        st.stop()


def append_by_headers(ws, row_dict, expected_headers):
  header_row = [str(h).strip() for h in ws.row_values(1)]
  if not header_row:
    header_row = expected_headers
    ws.append_row(header_row)

  missing = [h for h in expected_headers if h not in header_row]
  if missing:
    header_row = header_row + missing
    ws.update(range_name="A1", values=[header_row])

  aligned_row = [row_dict.get(col, "") for col in header_row]
  ws.append_row(aligned_row, value_input_option="USER_ENTERED")


def read_df(ws, headers, retries=3, delay=1.0):
  rows = None
  last_exc = None

  for attempt in range(retries):
    try:
      rows = ws.get_all_values()
      break
    except gspread.exceptions.APIError as exc:
      last_exc = exc
      if attempt < retries - 1:
        time.sleep(delay * (2**attempt))
      else:
        st.error(
            f"Error de Google Sheets al leer la pestaña '{ws.title}': {exc}"
        )
        st.stop()
    except Exception as exc:
      last_exc = exc
      if attempt < retries - 1:
        time.sleep(delay * (2**attempt))
      else:
        st.error(f"Error inesperado leyendo '{ws.title}': {last_exc}")
        st.stop()

  if not rows or len(rows) <= 1:
    return pd.DataFrame(columns=headers)

  raw_headers = [str(h).strip() for h in rows[0]]
  data_rows = rows[1:]

  safe_cols = [
      name if name else f"col_{idx}" for idx, name in enumerate(raw_headers)
  ]
  normalized_rows = []
  for r in data_rows:
    if not any(str(c).strip() != "" for c in r):
      continue
    padded = r[: len(safe_cols)] + [""] * max(0, len(safe_cols) - len(r))
    normalized_rows.append(dict(zip(safe_cols, padded)))

  if not normalized_rows:
    return pd.DataFrame(columns=headers)

  df = pd.DataFrame(normalized_rows)
  for col in headers:
    if col not in df.columns:
      df[col] = ""
  return df[headers]


def get_col_index(ws, col_name, fallback_idx):
  try:
    header_row = [str(h).strip() for h in ws.row_values(1)]
    if col_name in header_row:
      return header_row.index(col_name) + 1
  except Exception:
    pass
  return fallback_idx


def parse_num(series):
  return pd.to_numeric(
      series.astype(str)
      .str.replace("$", "", regex=False)
      .str.replace(",", "", regex=False)
      .str.strip(),
      errors="coerce",
  ).fillna(0.0)


def es_pago_movil(metodo):
  m = str(metodo or "").strip().lower()
  return ("móvil" in m) or ("movil" in m) or ("transferencia" in m)


def append_movimiento(ws_mov, tipo, referencia, monto, detalle, fecha_ref):
  hora_actual = datetime.now().strftime("%H:%M:%S")
  append_by_headers(
      ws_mov,
      {
          "id_mov": f"MOV-{uuid.uuid4().hex[:8].upper()}",
          "tipo": tipo,
          "referencia": referencia,
          "monto": float(monto),
          "detalle": detalle,
          "fecha_hora": f"{fecha_ref} {hora_actual}",
      },
      COLS_MOVS,
  )


# --- Conexión y preparación de hojas ---
sh = get_spreadsheet()
ws_caja = ensure_ws(sh, "caja_inicial", COLS_CAJA)
ws_clientes = ensure_ws(sh, "clientes", COLS_CLIENTES)
ws_prestamos = ensure_ws(sh, "prestamos", COLS_PRESTAMOS)
ws_pagos = ensure_ws(sh, "historial_pagos", COLS_PAGOS)
ws_gastos = ensure_ws(sh, "historial_gastos", COLS_GASTOS)
ws_movs = ensure_ws(sh, "historial_movimientos", COLS_MOVS)

df_caja = read_df(ws_caja, COLS_CAJA)
df_clientes = read_df(ws_clientes, COLS_CLIENTES)
df_prestamos = read_df(ws_prestamos, COLS_PRESTAMOS)
df_pagos = read_df(ws_pagos, COLS_PAGOS)
df_gastos = read_df(ws_gastos, COLS_GASTOS)
df_movs = read_df(ws_movs, COLS_MOVS)

# --- Sidebar: Fecha operativa y Cuadre de Caja Inicial Único ---
with st.sidebar:
  st.header("Jornada y Caja Inicial")
  fecha_sel = st.date_input(
      "Fecha operativa (Día a día)",
      value=date.today(),
      key="input_fecha_operativa",
  )
  fecha_str = fecha_sel.strftime("%Y-%m-%d")

  cobrador_activo = st.text_input(
      "Nombre del cobrador",
      value="Cobrador 01",
      key="input_cobrador_activo",
  )

  caja_fecha = (
      df_caja[df_caja["fecha"].astype(str).str.strip() == fecha_str]
      if not df_caja.empty
      else pd.DataFrame()
  )

  if not caja_fecha.empty:
    monto_caja_inicial = float(
        parse_num(caja_fecha["monto_inicial"]).iloc[0]
    )
    st.success(
        f"Caja abierta ({fecha_str}): **${monto_caja_inicial:,.2f}**"
    )
    st.caption(
        "El cuadre inicial es único por fecha y no cambia al registrar clientes,"
        " préstamos ni gastos."
    )
  else:
    monto_caja_inicial = 0.0
    st.warning(
        f"Aún no has registrado el cuadre de caja inicial para {fecha_str}."
    )
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
        hora_actual = datetime.now().strftime("%H:%M:%S")
        append_by_headers(
            ws_caja,
            {
                "fecha": fecha_str,
                "monto_inicial": float(monto_input),
                "cobrador": nombre_cob,
                "estado": "ABIERTA",
                "timestamp": f"{fecha_str} {hora_actual}",
            },
            COLS_CAJA,
        )
        append_movimiento(
            ws_movs,
            "APERTURA_CAJA",
            fecha_str,
            monto_input,
            f"Cuadre inicial por {nombre_cob}",
            fecha_str,
        )
        st.rerun()

# --- Cálculos del día seleccionado ---
pagos_fecha = (
    df_pagos[df_pagos["fecha_hora"].astype(str).str.startswith(fecha_str)].copy()
    if not df_pagos.empty
    else pd.DataFrame(columns=COLS_PAGOS)
)

if not pagos_fecha.empty:
  pagos_fecha["monto_num"] = parse_num(pagos_fecha["monto_pagado"])
  es_movil_mask = pagos_fecha["metodo_pago"].apply(es_pago_movil)
  # Los pagos por Pago Móvil / Transferencia NO se suman al cuadre de caja
  total_cobrado_caja_fecha = pagos_fecha.loc[~es_movil_mask, "monto_num"].sum()
  total_cobrado_movil_fecha = pagos_fecha.loc[es_movil_mask, "monto_num"].sum()
else:
  total_cobrado_caja_fecha = 0.0
  total_cobrado_movil_fecha = 0.0

prestamos_fecha = (
    df_prestamos[
        df_prestamos["fecha_inicio"].astype(str).str.strip() == fecha_str
    ].copy()
    if not df_prestamos.empty
    else pd.DataFrame(columns=COLS_PRESTAMOS)
)

if not prestamos_fecha.empty:
  prestamos_fecha["monto_num"] = parse_num(prestamos_fecha["monto_prestado"])
  afecta_p_mask = (
      prestamos_fecha["afecta_caja"].astype(str).str.strip().str.upper() == "SI"
  )
  total_prestado_caja_fecha = prestamos_fecha.loc[
      afecta_p_mask, "monto_num"
  ].sum()
  total_prestado_fuera_caja_fecha = prestamos_fecha.loc[
      ~afecta_p_mask, "monto_num"
  ].sum()
else:
  total_prestado_caja_fecha = 0.0
  total_prestado_fuera_caja_fecha = 0.0

gastos_fecha = (
    df_gastos[
        df_gastos["fecha_hora"].astype(str).str.startswith(fecha_str)
    ].copy()
    if not df_gastos.empty
    else pd.DataFrame(columns=COLS_GASTOS)
)

if not gastos_fecha.empty:
  gastos_fecha["monto_num"] = parse_num(gastos_fecha["monto"])
  afecta_g_mask = (
      gastos_fecha["afecta_caja"]
      .replace("", "SI")
      .fillna("SI")
      .astype(str)
      .str.strip()
      .str.upper()
      == "SI"
  )
  gasto_no_movil_mask = ~gastos_fecha["metodo_pago"].apply(es_pago_movil)
  total_gastos_caja_fecha = gastos_fecha.loc[
      afecta_g_mask & gasto_no_movil_mask, "monto_num"
  ].sum()
  total_gastos_fuera_caja_fecha = gastos_fecha.loc[
      ~(afecta_g_mask & gasto_no_movil_mask), "monto_num"
  ].sum()
else:
  total_gastos_caja_fecha = 0.0
  total_gastos_fuera_caja_fecha = 0.0

efectivo_esperado_caja = (
    monto_caja_inicial
    + total_cobrado_caja_fecha
    - total_prestado_caja_fecha
    - total_gastos_caja_fecha
)
cartera_activa = (
    parse_num(df_prestamos["saldo_pendiente"]).sum()
    if not df_prestamos.empty
    else 0.0
)

# --- Encabezado y KPIs del día ---
st.title(f"Sistema de Cobro — Jornada {fecha_str}")

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Caja Inicial", f"${monto_caja_inicial:,.2f}")
c2.metric("Cobrado en Caja (Efec.)", f"${total_cobrado_caja_fecha:,.2f}")
c3.metric("Prestado (En Caja)", f"${total_prestado_caja_fecha:,.2f}")
c4.metric("Gastos (En Caja)", f"${total_gastos_caja_fecha:,.2f}")
c5.metric("Efectivo Esperado en Caja", f"${efectivo_esperado_caja:,.2f}")

st.caption(
    f"Pago Móvil / Transf. del día (no suma a caja):"
    f" **${total_cobrado_movil_fecha:,.2f}** | Prestado sin caja:"
    f" **${total_prestado_fuera_caja_fecha:,.2f}** | Gastos sin caja:"
    f" **${total_gastos_fuera_caja_fecha:,.2f}** | Cartera activa:"
    f" **${cartera_activa:,.2f}**"
)

if caja_fecha.empty:
  st.info(
      f"Registra primero el **Cuadre de Caja Inicial** para **{fecha_str}**"
      " en la barra lateral para operar."
  )
  st.stop()

# --- Pestañas operativas ---
(
    tab_cobro,
    tab_clientes,
    tab_nuevo_p,
    tab_gastos,
    tab_cartera,
    tab_historial,
) = st.tabs([
    "Registrar Cobro / Ruta",
    "Registrar Clientes",
    "Nuevo Préstamo",
    "Gastos",
    "Cartera de Clientes",
    "Día a Día e Historial",
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
      df_p_calc["caja_limpio"] = (
          df_p_calc["afecta_caja"]
          .replace("", "SI")
          .fillna("SI")
          .astype(str)
          .str.strip()
          .str.upper()
      )

      activos = df_p_calc[
          (df_p_calc["saldo_num"] > 0) & (df_p_calc["estado_limpio"] != "PAGADO")
      ].copy()

      filtro_caja_cobro = st.radio(
          "Filtrar préstamos disponibles para cobro",
          [
              "Todos con saldo",
              "Solo NO descuentan caja (NO)",
              "Solo SI descuentan caja (SI)",
          ],
          horizontal=True,
          key="radio_filtro_caja_cobro",
      )

      if filtro_caja_cobro == "Solo NO descuentan caja (NO)":
        activos = activos[activos["caja_limpio"] == "NO"]
      elif filtro_caja_cobro == "Solo SI descuentan caja (SI)":
        activos = activos[activos["caja_limpio"] == "SI"]

      if activos.empty:
        st.info("No hay préstamos con saldo pendiente bajo este filtro.")
      else:
        activos["etiqueta"] = activos.apply(
            lambda r: (
                f"{r['id_prestamo']} | {r['cliente']} | Saldo por cobrar:"
                f" ${r['saldo_num']:,.2f} | Cuota: ${r['cuota_num']:,.2f} |"
                f" Caja: {r['caja_limpio']}"
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
            hora_actual = datetime.now().strftime("%H:%M:%S")
            fecha_hora_reg = f"{fecha_str} {hora_actual}"
            nombre_cob = cobrador_activo.strip() or "Cobrador"

            append_by_headers(
                ws_pagos,
                {
                    "id_pago": id_pago,
                    "id_prestamo": fila_p["id_prestamo"],
                    "cliente": fila_p["cliente"],
                    "monto_pagado": float(monto_pago),
                    "metodo_pago": metodo,
                    "saldo_restante": float(nuevo_saldo),
                    "fecha_hora": fecha_hora_reg,
                    "cobrador": nombre_cob,
                },
                COLS_PAGOS,
            )

            col_id_p = get_col_index(ws_prestamos, "id_prestamo", 1)
            cell = ws_prestamos.find(
                str(fila_p["id_prestamo"]), in_column=col_id_p
            )
            if cell:
              col_saldo = get_col_index(ws_prestamos, "saldo_pendiente", 10)
              col_estado = get_col_index(ws_prestamos, "estado", 13)
              ws_prestamos.update_cell(cell.row, col_saldo, float(nuevo_saldo))
              ws_prestamos.update_cell(cell.row, col_estado, nuevo_estado)

            nota_caja = (
                "NO suma a caja (Pago Móvil/Transf.)"
                if es_pago_movil(metodo)
                else "Suma a caja"
            )
            append_movimiento(
                ws_movs,
                "COBRO",
                id_pago,
                monto_pago,
                f"Abono a {fila_p['id_prestamo']} ({fila_p['cliente']}) |"
                f" {metodo} ({nota_caja}) | Saldo: ${nuevo_saldo:,.2f}",
                fecha_str,
            )
            st.success(
                f"Pago {id_pago} registrado ({metodo}). Saldo restante:"
                f" ${nuevo_saldo:,.2f}"
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
        append_by_headers(
            ws_clientes,
            {
                "id_cliente": id_cliente,
                "nombre": cli_nombre.strip(),
                "telefono": cli_telefono.strip(),
                "direccion": cli_direccion.strip(),
                "referencia": cli_referencia.strip(),
                "fecha_registro": fecha_str,
                "estado": "ACTIVO",
            },
            COLS_CLIENTES,
        )
        append_movimiento(
            ws_movs,
            "NUEVO_CLIENTE",
            id_cliente,
            0.0,
            f"Cliente registrado: {cli_nombre.strip()} ({cli_telefono.strip()})",
            fecha_str,
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

        append_by_headers(
            ws_prestamos,
            {
                "id_prestamo": id_prestamo,
                "id_cliente": cli_fila["id_cliente"],
                "cliente": cli_fila["nombre"],
                "telefono": cli_fila["telefono"],
                "monto_prestado": float(monto_prestado),
                "interes_pct": float(interes_pct),
                "total_a_pagar": float(total_pagar),
                "cuota_valor": float(cuota_val),
                "frecuencia": frecuencia,
                "saldo_pendiente": float(total_pagar),
                "fecha_inicio": fecha_str,
                "afecta_caja": flag_caja,
                "estado": "ACTIVO",
            },
            COLS_PRESTAMOS,
        )
        append_movimiento(
            ws_movs,
            "NUEVO_PRESTAMO",
            id_prestamo,
            monto_prestado,
            f"Cliente: {cli_fila['nombre']} ({cli_fila['id_cliente']}) |"
            f" Total: ${total_pagar:,.2f} | Afecta caja: {flag_caja}",
            fecha_str,
        )
        st.success(
            f"Préstamo {id_prestamo} creado para **{cli_fila['nombre']}**."
            f" Saldo disponible por cobrar: **${total_pagar:,.2f}** | Afecta"
            f" caja: **{flag_caja}**"
        )
        st.rerun()

with tab_gastos:
  st.subheader("Registrar Gasto Operativo")
  with st.form("form_nuevo_gasto", clear_on_submit=True):
    g_col1, g_col2, g_col3 = st.columns(3)
    gasto_cat = g_col1.selectbox(
        "Categoría",
        [
            "Gasolina / Transporte",
            "Comida / Viáticos",
            "Recargas / Datos",
            "Mantenimiento / Repuestos",
            "Otros",
        ],
        key="select_gasto_categoria",
    )
    gasto_monto = g_col2.number_input(
        "Monto del gasto ($)",
        min_value=0.01,
        value=5.0,
        step=1.0,
        format="%.2f",
        key="input_gasto_monto",
    )
    gasto_metodo = g_col3.selectbox(
        "Método de pago",
        ["Efectivo", "Pago Móvil / Transferencia", "Otro"],
        key="select_gasto_metodo",
    )

    gasto_desc = st.text_input(
        "Detalle o concepto del gasto *",
        placeholder="Ej. Combustible moto ruta norte",
        key="input_gasto_desc",
    )

    descontar_gasto_caja = st.toggle(
        "Descontar este gasto de la caja del día (SI / NO)",
        value=True,
        key="toggle_descontar_gasto_caja",
    )

    guardar_gasto = st.form_submit_button("Registrar Gasto")
    if guardar_gasto:
      if not gasto_desc.strip():
        st.error("Ingresa un detalle o concepto para el gasto.")
      else:
        id_gasto = f"GAS-{uuid.uuid4().hex[:8].upper()}"
        flag_gasto_caja = "SI" if descontar_gasto_caja else "NO"
        hora_actual = datetime.now().strftime("%H:%M:%S")
        fecha_hora_reg = f"{fecha_str} {hora_actual}"
        nombre_cob = cobrador_activo.strip() or "Cobrador"

        append_by_headers(
            ws_gastos,
            {
                "id_gasto": id_gasto,
                "categoria": gasto_cat,
                "descripcion": gasto_desc.strip(),
                "monto": float(gasto_monto),
                "metodo_pago": gasto_metodo,
                "afecta_caja": flag_gasto_caja,
                "fecha_hora": fecha_hora_reg,
                "cobrador": nombre_cob,
            },
            COLS_GASTOS,
        )
        append_movimiento(
            ws_movs,
            "GASTO",
            id_gasto,
            gasto_monto,
            f"[{gasto_cat}] {gasto_desc.strip()} | {gasto_metodo} | Afecta"
            f" caja: {flag_gasto_caja}",
            fecha_str,
        )
        st.success(
            f"Gasto {id_gasto} por **${float(gasto_monto):,.2f}** registrado"
            f" ({gasto_metodo} | Afecta caja: **{flag_gasto_caja}**)."
        )
        st.rerun()

  st.markdown(f"### Gastos Registrados ({fecha_str})")
  if gastos_fecha.empty:
    st.info(f"No se han registrado gastos en la fecha {fecha_str}.")
  else:
    st.dataframe(
        gastos_fecha[COLS_GASTOS], use_container_width=True, hide_index=True
    )

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
  st.subheader("Resumen Día a Día por Fecha")

  fechas_set = set()
  if not df_caja.empty:
    fechas_set.update(
        [d.strip() for d in df_caja["fecha"].astype(str) if d.strip()]
    )
  if not df_prestamos.empty:
    fechas_set.update([
        d.strip() for d in df_prestamos["fecha_inicio"].astype(str) if d.strip()
    ])
  if not df_pagos.empty:
    fechas_set.update([
        d.astype(str).str[:10].strip()
        for _, d in df_pagos["fecha_hora"].items()
        if str(d).strip()
    ])
  if not df_gastos.empty:
    fechas_set.update([
        d.astype(str).str[:10].strip()
        for _, d in df_gastos["fecha_hora"].items()
        if str(d).strip()
    ])

  fechas_ordenadas = sorted(fechas_set, reverse=True)
  resumen_dias = []

  for f_dia in fechas_ordenadas:
    c_dia = (
        df_caja[df_caja["fecha"].astype(str).str.strip() == f_dia]
        if not df_caja.empty
        else pd.DataFrame()
    )
    m_ini = (
        float(parse_num(c_dia["monto_inicial"]).iloc[0])
        if not c_dia.empty
        else 0.0
    )

    p_dia = (
        df_pagos[df_pagos["fecha_hora"].astype(str).str.startswith(f_dia)]
        if not df_pagos.empty
        else pd.DataFrame(columns=COLS_PAGOS)
    )
    if not p_dia.empty:
      p_vals = parse_num(p_dia["monto_pagado"])
      p_mov_mask = p_dia["metodo_pago"].apply(es_pago_movil)
      cob_caja = float(p_vals[~p_mov_mask].sum())
      cob_movil = float(p_vals[p_mov_mask].sum())
    else:
      cob_caja = 0.0
      cob_movil = 0.0

    pr_dia = (
        df_prestamos[df_prestamos["fecha_inicio"].astype(str).str.strip() == f_dia]
        if not df_prestamos.empty
        else pd.DataFrame(columns=COLS_PRESTAMOS)
    )
    if not pr_dia.empty:
      pr_vals = parse_num(pr_dia["monto_prestado"])
      pr_caja_mask = (
          pr_dia["afecta_caja"].astype(str).str.strip().str.upper() == "SI"
      )
      pres_caja = float(pr_vals[pr_caja_mask].sum())
    else:
      pres_caja = 0.0

    g_dia = (
        df_gastos[df_gastos["fecha_hora"].astype(str).str.startswith(f_dia)]
        if not df_gastos.empty
        else pd.DataFrame(columns=COLS_GASTOS)
    )
    if not g_dia.empty:
      g_vals = parse_num(g_dia["monto"])
      g_caja_mask = (
          g_dia["afecta_caja"]
          .replace("", "SI")
          .fillna("SI")
          .astype(str)
          .str.strip()
          .str.upper()
          == "SI"
      ) & (~g_dia["metodo_pago"].apply(es_pago_movil))
      gas_caja = float(g_vals[g_caja_mask].sum())
    else:
      gas_caja = 0.0

    efec_cierre = round(m_ini + cob_caja - pres_caja - gas_caja, 2)
    resumen_dias.append({
        "fecha": f_dia,
        "caja_inicial": round(m_ini, 2),
        "cobrado_caja": round(cob_caja, 2),
        "cobrado_pago_movil": round(cob_movil, 2),
        "prestado_caja": round(pres_caja, 2),
        "gastos_caja": round(gas_caja, 2),
        "efectivo_esperado_caja": efec_cierre,
    })

  if resumen_dias:
    st.dataframe(
        pd.DataFrame(resumen_dias), use_container_width=True, hide_index=True
    )
  else:
    st.info("Aún no hay movimientos registrados por fecha.")

  st.subheader("Historial de Pagos (Solo Lectura / Nunca se Borra)")
  st.dataframe(df_pagos, use_container_width=True, hide_index=True)

  st.subheader("Historial de Gastos (Solo Lectura / Nunca se Borra)")
  st.dataframe(df_gastos, use_container_width=True, hide_index=True)

  st.subheader("Bitácora de Movimientos del Sistema")
  st.dataframe(df_movs, use_container_width=True, hide_index=True)
