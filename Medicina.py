import datetime
import io
import pathlib
import re
from PIL import Image as PILImage, ImageDraw
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st

st.set_page_config(page_title="Servicio de Medicina - Registro HIS", layout="wide")

if "lista_pacientes_medicina" not in st.session_state:
  st.session_state.lista_pacientes_medicina = []

# --- REINICIO SEGURO DE ESTADOS (PACIENTES Y DIAGNÓSTICOS) ANTES DE RENDERIZAR ---
if "reset_form" in st.session_state and st.session_state.reset_form:
  st.session_state.form_dni = ""
  st.session_state.input_nombres = ""
  st.session_state.input_edad = 30
  st.session_state.input_sexo = "Seleccione..."
  # Limpiar diagnósticos, campos manuales y laboratorios anteriores  
  for i in range(1, 6):
    st.session_state[f"diag_{i}"] = "-- Seleccione o busque CIE-10 --"
    st.session_state[f"diag_manual_{i}"] = ""
    st.session_state[f"lab_{i}"] = ""
    st.session_state[f"prd_{i}"] = "D"
  # Limpiar antropometría opcional
  st.session_state.input_talla = ""
  st.session_state.input_peso = ""
  st.session_state.input_pab = ""
  st.session_state.input_hb = ""
  st.session_state.reset_form = False


# --- CARGA MASIVA AUTOMÁTICA DEL CATÁLOGO CIE-10 OFICIAL DESDE EXCEL ---
@st.cache_data
def cargar_catalogo_cie10():
  catalogo = {}
  try:
    script_dir = (
        pathlib.Path(__file__).parent
        if "__file__" in locals()
        else pathlib.Path.cwd()
    )
    archivo_excel = script_dir / "cie10_minsa.xlsx"

    df = pd.read_excel(archivo_excel, sheet_name=0, header=None, dtype=str)
    patron = re.compile(
        r"^([A-Z]\d{2,3}(?:\.\d)?|\d{4,5}(?:\.\d{1,2})?)\s*[-–]\s*(.+)",
        re.IGNORECASE,
    )

    for columna in df.columns:
      for valor in df[columna].dropna():
        texto = str(valor).strip()
        match = patron.match(texto)
        if match:
          codigo = match.group(1).strip().upper()
          descripcion = match.group(2).strip().upper()
          catalogo[codigo] = descripcion

  except Exception as e:
    print(f"Error al leer el archivo Excel oficial CIE-10: {e}")

  if not catalogo:
    catalogo = {
        "J029": "FARINGITIS AGUDA, NO ESPECIFICADA",
        "J039": "AMIGDALITIS AGUDA, NO ESPECIFICADA",
    }

  return catalogo


CIE10_DICT = cargar_catalogo_cie10()
opciones_cie10 = [
    "-- Seleccione o busque CIE-10 --",
    "-- ✏️ Ingresar CIE-10 Manualmente --",
] + [f"{cod} - {desc}" for cod, desc in CIE10_DICT.items()]


# --- CARGA DEL PADRÓN DE PACIENTES (NUMDOC, PACIENTE, FECHANAC, SEXO) ---
@st.cache_data
def cargar_pacientes():
  pacientes_dict = {}
  try:
    script_dir = (
        pathlib.Path(__file__).parent
        if "__file__" in locals()
        else pathlib.Path.cwd()
    )
    archivo_excel = script_dir / "pacientes_minsa.xlsx"

    df = pd.read_excel(archivo_excel, sheet_name=0, dtype=str)

    if {"NUMDOC", "PACIENTE", "FECHANAC", "SEXO"}.issubset(df.columns):
      for _, row in df.dropna(subset=["NUMDOC"]).iterrows():
        dni = str(row["NUMDOC"]).strip().replace(".0", "")
        pacientes_dict[dni] = {
            "nombres": str(row["PACIENTE"]).strip().upper(),
            "fechanac": str(row["FECHANAC"]).strip(),
            "sexo": str(row["SEXO"]).strip().upper(),
        }
  except Exception as e:
    print(f"Error al leer el padrón de pacientes: {e}")

  return pacientes_dict


PACIENTES_DICT = cargar_pacientes()

medicos_predefinidos = {
    "Seleccione DNI...": {"nombres": ""},
    "71981216": {"nombres": "ESCALAYA CORDOVA, GIANCARLOS JESUS"},
    "42568879": {"nombres": "GUILLERMO COYCCA, MICHAEL JACKSON"},
    "46007585": {"nombres": "HUARCAYA COTAQUISPE, NOELI"},
    "28291920": {"nombres": "CORDOVA PURE, MICHAEL"},
    "70190930": {"nombres": "ALFARO MENDIVEL, FRANK"},
    "42753517": {"nombres": "YUPANQUI VARGAS, EVELIN"},
    "74248771": {"nombres": "MEDRANO TORRES, ALDO DAVID"},
    "28290282": {"nombres": "PRETELL NACCHA, ORLANDO"},
    "47620489": {"nombres": "ALBUJAR ONOFRE, LUIS ANGEL"},
    "05982514": {"nombres": "ARCE SUBELZA, ELIAS MOISES"},
    "05351943": {"nombres": "ALCOBA SANCHEZ, WILLIAM"},
    "70445685": {"nombres": "FERNANDEZ CURI, IVETTE CARMEN"},
    "72743728": {"nombres": "VASQUEZ DE LA CRUZ, ALEXIS WALTER"},
    "72936147": {"nombres": "SARAVIA TASAYCO, RITA VALERIA"},
    "41817213": {"nombres": "ESPEZA PALACIOS, EDITH KETTY"},
    "08230872": {"nombres": "TOLABA RAMOS, MARIA ELENA"},
    "47129450": {"nombres": "PEREYRA GONZALES, ALINA KATHERINE"},
    "73898542": {"nombres": "MUCHARI VELAPATIÑO, BRYAN STEVE"},
    "MANUAL": {"nombres": "MANUAL"},
}

st.markdown("### MICRORED BELEN")
st.markdown("### 🏥 Servicio de Medicina")
st.markdown("**Registro Diario de Atención - Formato HIS MINSA**")

st.markdown("#### 1. Información del Establecimiento y Profesional")
c1, c2, c3, c4, c5, c6 = st.columns([1.1, 1.3, 1.4, 1.4, 1.1, 2.2])

with c1:
  anio = st.text_input("Año *", "2026")
with c2:
  mes = st.selectbox(
      "Mes *",
      [
          "Enero",
          "Febrero",
          "Marzo",
          "Abril",
          "Mayo",
          "Junio",
          "Julio",
          "Agosto",
          "Septiembre",
          "Octubre",
          "Noviembre",
          "Diciembre",
      ],
      index=8,
  )
with c3:
  centro_salud = st.selectbox(
      "Centro de Salud *",
      [
          "Seleccione un Centro...",
          "Belén",
          "Barrios Altos",
          "Huascahura",
          "Morro de Arica",
          "Rancha",
          "Santa Ana",
      ],
  )
with c4:
  dni_seleccionado = st.selectbox(
      "DNI Profesional *", list(medicos_predefinidos.keys())
  )
with c5:
  turno_op = st.selectbox(
      "Turno *", ["Mañana (M)", "Tarde (T)", "Noche (N)"]
  )
with c6:
  ups_opcion = st.selectbox(
      "UPS / Servicio *",
      [
          "302303 - MEDICINA GENERAL",
          "300303 - CIRUGIA EN CONS. EXT / TOPICO",
          "301604 - ECOGRAFIA GINECO-OBSTETRICA",
      ],
  )

datos_med = medicos_predefinidos[dni_seleccionado]

if dni_seleccionado == "MANUAL":
  c_m1, c_m2 = st.columns(2)
  with c_m1:
    dni_profesional = st.text_input(
        "DNI Manual (8 dígitos)", max_chars=8, value=""
    )
  with c_m2:
    nombres_profesional = st.text_input(
        "Nombres del Profesional Médico",
        value="",
        placeholder="Apellidos y Nombres",
    )
elif dni_seleccionado == "Seleccione DNI...":
  dni_profesional = ""
  nombres_profesional = ""
  st.text_input(
      "Nombres del Profesional Médico (Automático)", value="", disabled=True
  )
else:
  dni_profesional = dni_seleccionado
  nombres_profesional = datos_med["nombres"]
  st.text_input(
      "Nombres del Profesional Médico (Automático)",
      value=nombres_profesional,
      disabled=True,
  )

st.markdown("---")

st.markdown(
    f"#### 2. Identificación del Paciente y Diagnósticos (Nuevo Paciente N°"
    f" {len(st.session_state.lista_pacientes_medicina) + 1})"
)

# --- INICIALIZACIÓN DE ESTADOS PARA EL PACIENTE ---
if "form_dni" not in st.session_state:
  st.session_state.form_dni = ""
if "input_nombres" not in st.session_state:
  st.session_state.input_nombres = ""
if "input_edad" not in st.session_state:
  st.session_state.input_edad = 30
if "input_sexo" not in st.session_state:
  st.session_state.input_sexo = "Seleccione..."


def actualizar_paciente_por_dni():
  dni_input = st.session_state.get("form_dni", "").strip()
  if dni_input in PACIENTES_DICT:
    p_info = PACIENTES_DICT[dni_input]
    st.session_state.input_nombres = p_info["nombres"]
    st.session_state.input_sexo = (
        "F (Femenino)" if p_info["sexo"] == "F" else "M (Masculino)"
    )

    try:
      fn = pd.to_datetime(p_info["fechanac"], errors="coerce")
      if pd.notna(fn):
        ref_date = datetime.date.today()
        anos = (
            ref_date.year
            - fn.year
            - ((ref_date.month, ref_date.day) < (fn.month, fn.day))
        )
        st.session_state.input_edad = max(0, int(anos))
    except Exception:
      pass
  else:
    st.session_state.input_nombres = ""
    st.session_state.input_edad = 30
    st.session_state.input_sexo = "Seleccione..."


col_search_1, col_search_2 = st.columns([1, 3])
with col_search_1:
  dni_paciente = st.text_input(
      "DNI del Paciente *",
      max_chars=8,
      placeholder="Escribe DNI",
      key="form_dni",
      on_change=actualizar_paciente_por_dni,
  )
with col_search_2:
  if dni_paciente and dni_paciente in PACIENTES_DICT:
    st.success("¡Paciente encontrado en el Padrón! Datos autocompletados. ✔️")
  elif dni_paciente:
    st.info("DNI no registrado en el padrón. Complete los datos manualmente.")

# --- CAMPOS DE PACIENTE ---
cp1, cp2, cp3, cp4 = st.columns(4)
with cp1:
  fecha_atencion = st.date_input("Fecha de Atención *", datetime.date.today())
with cp2:
  nombres_paciente = st.text_input(
      "Nombres del Paciente *",
      placeholder="APELLIDOS Y NOMBRES",
      key="input_nombres",
  )
with cp3:
  edad_anos = st.number_input(
      "Edad *", min_value=0, max_value=120, key="input_edad"
  )
with cp4:
  sexo_paciente = st.selectbox(
      "Sexo *", ["Seleccione...", "F (Femenino)", "M (Masculino)"], key="input_sexo"
  )

condicion_paciente = st.radio(
    "Condición Paciente *",
    ["Nuevo (N)", "Continuador (C)", "Reingresante (R)"],
    horizontal=True,
    key="input_condicion",
)

st.markdown("---")
st.markdown(
    "#### Diagnósticos del Servicio de Medicina (Búsqueda y Autocompletado"
    " CIE-10)"
)

hc1, hc2, hc3 = st.columns([1, 6, 1])
with hc1:
  st.markdown("**N°**")
with hc2:
  st.markdown("**SELECCIONE O BUSQUE DIAGNÓSTICO CIE-10**")
with hc3:
  st.markdown("**ACCIONES**")

for i in range(1, 6):
  r1, r2, r3, r4 = st.columns([1, 5, 1, 1])
  with r1:
    st.markdown(f"#### {i}")
  with r2:
    seleccion_actual = st.session_state.get(
        f"diag_{i}", "-- Seleccione o busque CIE-10 --"
    )
    if seleccion_actual == "-- ✏️ Ingresar CIE-10 Manualmente --":
      sub_c1, sub_c2 = st.columns([2, 3])
      with sub_c1:
        st.selectbox(
            f"Diagnóstico {i}",
            opciones_cie10,
            key=f"diag_{i}",
            label_visibility="collapsed",
        )
      with sub_c2:
        st.text_input(
            f"Manual {i}",
            placeholder="Ej: U07.2 - DESCRIPCION",
            key=f"diag_manual_{i}",
            label_visibility="collapsed",
        )
    else:
      st.selectbox(
          f"Diagnóstico {i}",
          opciones_cie10,
          key=f"diag_{i}",
          label_visibility="collapsed",
      )
  with r3:
    st.selectbox(
        "P/D/R", ["P", "D", "R"], key=f"prd_{i}", label_visibility="collapsed"
    )
  with r4:
    st.text_input(
        "Lab",
        value="",
        key=f"lab_{i}",
        label_visibility="collapsed",
        placeholder="LAB",
    )

st.markdown("---")

with st.expander("📊 Datos Antropométricos Opcionales (Talla, Peso, Pab, Hb)"):
  ca1, ca2, ca3, ca4 = st.columns(4)
  with ca1:
    talla_str = st.text_input(
        "Talla (cm)", value="", placeholder="Ej: 165.0", key="input_talla"
    )
  with ca2:
    peso_str = st.text_input(
        "Peso (kg)", value="", placeholder="Ej: 70.0", key="input_peso"
    )
  with ca3:
    pab_str = st.text_input(
        "Pab (cm)", value="", placeholder="Ej: 85.0", key="input_pab"
    )
  with ca4:
    hb_str = st.text_input(
        "Hemoglobina", value="", placeholder="Ej: 14.0", key="input_hb"
    )

btn_guardar = st.button("➕ Guardar e Ingresar Siguiente", type="primary")

if btn_guardar:
  dni_pac_limpio = re.sub(r"\D", "", dni_paciente)
  dni_prof_limpio = re.sub(r"\D", "", dni_profesional)

  codigos_ingresados = []
  tiene_duplicados = False

  for i in range(1, 6):
    seleccion = st.session_state.get(f"diag_{i}", "")
    if seleccion == "-- ✏️ Ingresar CIE-10 Manualmente --":
      seleccion = st.session_state.get(f"diag_manual_{i}", "").strip()

    if seleccion and seleccion != "-- Seleccione o busque CIE-10 --":
      if " - " in seleccion:
        codigo = seleccion.split(" - ", 1)[0].strip()
      else:
        partes_man = seleccion.split(" ", 1)
        codigo = partes_man[0].strip() if partes_man else seleccion

      if codigo in codigos_ingresados:
        tiene_duplicados = True
        break
      codigos_ingresados.append(codigo)

  if not centro_salud or centro_salud == "Seleccione un Centro...":
    st.error("⚠️ Por favor, seleccione un Centro de Salud.")
  elif not dni_seleccionado or dni_seleccionado == "Seleccione DNI...":
    st.error("⚠️ Por favor, seleccione el DNI del Profesional.")
  elif not dni_prof_limpio or len(dni_prof_limpio) != 8:
    st.error("⚠️ El DNI del Profesional debe contener exactamente 8 dígitos.")
  elif not dni_pac_limpio or len(dni_pac_limpio) != 8:
    st.error("⚠️ El DNI del Paciente debe contener exactamente 8 dígitos.")
  elif not nombres_paciente:
    st.error("⚠️ Por favor, ingrese los nombres del paciente.")
  elif sexo_paciente == "Seleccione...":
    st.error("⚠️ Por favor, seleccione el sexo del paciente.")
  elif tiene_duplicados:
    st.error(
        "⚠️ No se permiten diagnósticos duplicados en una misma atención. Por"
        " favor verifique los CIE-10 seleccionados."
    )
  else:
    cond_letra = (
        "N"
        if "Nuevo" in condicion_paciente
        else ("C" if "Continuador" in condicion_paciente else "R")
    )
    sexo_letra = sexo_paciente[0]

    val_talla = talla_str.strip() if talla_str.strip() != "" else ""
    val_peso = peso_str.strip() if peso_str.strip() != "" else ""
    val_pab = pab_str.strip() if pab_str.strip() != "" else ""
    val_hb = hb_str.strip() if hb_str.strip() != "" else ""

    partes = []
    if val_talla:
      partes.append("Talla: " + val_talla)
    if val_peso:
      partes.append("Peso: " + val_peso)
    if val_pab:
      partes.append("P. Abd: " + val_pab)
    if val_hb:
      partes.append("Hb: " + val_hb)
    medidas_text = " / ".join(partes) if partes else "-"

    codigos_list = []
    descripciones_list = []
    labs_list = []
    prds_list = []

    for i in range(1, 6):
      seleccion = st.session_state.get(f"diag_{i}", "")
      if seleccion == "-- ✏️ Ingresar CIE-10 Manualmente --":
        seleccion = st.session_state.get(f"diag_manual_{i}", "").strip()

      l = st.session_state.get(f"lab_{i}", "").strip()
      pr = st.session_state.get(f"prd_{i}", "D")

      if seleccion and seleccion != "-- Seleccione o busque CIE-10 --":
        if " - " in seleccion:
          partes_diag = seleccion.split(" - ", 1)
          c = partes_diag[0].strip()
          d = partes_diag[1].strip() if len(partes_diag) > 1 else "-"
        else:
          partes_diag = seleccion.split(" ", 1)
          c = partes_diag[0].strip() if partes_diag else seleccion
          d = partes_diag[1].strip() if len(partes_diag) > 1 else "-"

        codigos_list.append(c)
        descripciones_list.append(d)
        labs_list.append(l)
        prds_list.append(pr)

    while len(codigos_list) < 3:
      codigos_list.append("")
      descripciones_list.append("-")
      labs_list.append("")
      prds_list.append("")

    nuevo_paciente = {
        "nro": len(st.session_state.lista_pacientes_medicina) + 1,
        "dni": dni_pac_limpio,
        "nombres": nombres_paciente,
        "edad": edad_anos,
        "sexo": sexo_letra,
        "dia_atencion": fecha_atencion.strftime("%d"),
        "antropometria": medidas_text,
        "condicion": cond_letra,
        "codigos": codigos_list,
        "descripciones": descripciones_list,
        "labs": labs_list,
        "prds": prds_list,
    }
    st.session_state.lista_pacientes_medicina.append(nuevo_paciente)

    st.session_state.reset_form = True

    st.success(
        "✅ ¡Paciente guardado con éxito! (Total:"
        f" {len(st.session_state.lista_pacientes_medicina)})"
    )
    st.rerun()

st.markdown("---")
st.markdown(
    f"### 📋 Panel de Pacientes Registrados"
    f" ({len(st.session_state.lista_pacientes_medicina)})"
)

if len(st.session_state.lista_pacientes_medicina) > 0:
  df_preview = pd.DataFrame([
      {
          "nro": p["nro"],
          "dia": p["dia_atencion"],
          "dni": p["dni"],
          "nombres": p["nombres"],
          "edad": p["edad"],
          "sexo": p["sexo"],
          "antropometria": p["antropometria"],
          "condicion": p["condicion"],
          "codigos": " | ".join([c for c in p["codigos"] if c != ""]),
          "diagnosticos": " | ".join(
              [d for d in p["descripciones"] if d != "-"]
          ),
          "labs": " | ".join([l for l in p["labs"] if l != ""]),
      }
      for p in st.session_state.lista_pacientes_medicina
  ])
  st.dataframe(df_preview, use_container_width=True)
else:
  st.info("ℹ️ Aún no hay pacientes guardados.")

with st.sidebar:
  st.markdown("### 📥 Generar Reporte PDF")
  st.info(
      f"Pacientes en lista:"
      f" **{len(st.session_state.lista_pacientes_medicina)}**"
  )

  if len(st.session_state.lista_pacientes_medicina) > 0:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=portrait(A4),
        rightMargin=15,
        leftMargin=15,
        topMargin=15,
        bottomMargin=15,
    )
    styles = getSampleStyleSheet()

    cell_style = ParagraphStyle(
        "CellStyle",
        parent=styles["Normal"],
        fontSize=6,
        leading=7.5,
        textColor=colors.black,
    )
    cell_center = ParagraphStyle(
        "CellCenter",
        parent=styles["Normal"],
        fontSize=6,
        leading=7.5,
        alignment=1,
        textColor=colors.black,
    )
    cell_edad = ParagraphStyle(
        "CellEdad",
        parent=styles["Normal"],
        fontSize=5.5,
        leading=6.5,
        alignment=1,
        textColor=colors.black,
    )
    header_style = ParagraphStyle(
        "HeaderStyle",
        parent=styles["Normal"],
        fontSize=6.5,
        leading=8,
        alignment=1,
        fontName="Helvetica-Bold",
        textColor=colors.black,
    )

    title_main = ParagraphStyle(
        "TitleMain",
        parent=styles["Normal"],
        fontSize=8.5,
        leading=10,
        alignment=1,
        fontName="Helvetica-Bold",
        textColor=colors.black,
    )
    title_sub = ParagraphStyle(
        "TitleSub",
        parent=styles["Normal"],
        fontSize=6.5,
        leading=8,
        alignment=1,
        fontName="Helvetica-Bold",
        textColor=colors.black,
    )
    title_desc = ParagraphStyle(
        "TitleDesc",
        parent=styles["Normal"],
        fontSize=6.5,
        leading=8,
        alignment=1,
        fontName="Helvetica-Bold",
        textColor=colors.black,
    )
    lbl_style = ParagraphStyle(
        "LblStyle",
        parent=styles["Normal"],
        fontSize=5,
        leading=6,
        alignment=0,
        fontName="Helvetica-Bold",
        textColor=colors.black,
    )
    val_style = ParagraphStyle(
        "ValStyle",
        parent=styles["Normal"],
        fontSize=6,
        leading=7.5,
        alignment=1,
        textColor=colors.black,
    )

    elements = []

    m_check = "X" if "Mañana" in turno_op else ""
    t_check = "X" if "Tarde" in turno_op else ""
    n_check = "X" if "Noche" in turno_op else ""

    script_dir = (
        pathlib.Path(__file__).parent
        if "__file__" in locals()
        else pathlib.Path.cwd()
    )
    carpeta_firmas = script_dir / "firmas"

    img_element = None
    if dni_profesional:
      archivos_encontrados = list(carpeta_firmas.glob(f"{dni_profesional}.*"))
      if archivos_encontrados:
        try:
          ruta_img = archivos_encontrados[0]
          pil_img = PILImage.open(ruta_img).convert("RGB")
          img_io = io.BytesIO()
          pil_img.save(img_io, format="JPEG")
          img_io.seek(0)
          img_element = Image(img_io, width=90, height=50)
        except Exception:
          img_element = None

    if img_element:
      tabla_contenido_firma = Table(
          [
              [Paragraph("FIRMA Y SELLO RESPONSABLE HIS MINSA", lbl_style)],
              [Spacer(1, 2)],
              [img_element],
          ],
          colWidths=[160],
      )
      tabla_contenido_firma.setStyle(
          TableStyle([
              ("ALIGN", (0, 0), (-1, -1), "CENTER"),
              ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
              ("LEFTPADDING", (0, 0), (-1, -1), 2),
              ("RIGHTPADDING", (0, 0), (-1, -1), 2),
              ("TOPPADDING", (0, 0), (-1, -1), 2),
              ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
          ])
      )
      contenido_firma_celda = tabla_contenido_firma
    else:
      contenido_firma_celda = [
          Paragraph("FIRMA Y SELLO RESPONSABLE HIS MINSA", lbl_style),
          Spacer(1, 15),
          Paragraph("", cell_style),
      ]

    col_izquierda_data = [
        [Paragraph("LOTE", lbl_style), Paragraph("", val_style)],
        [Paragraph("PAGINA", lbl_style), Paragraph("1", val_style)],
        [Paragraph("FECHA PROCES.", lbl_style), Paragraph("", val_style)],
        [Paragraph("DNI DIGITADOR", lbl_style), Paragraph("", val_style)],
    ]
    t_col_izquierda = Table(
        col_izquierda_data,
        colWidths=[55, 45],
        rowHeights=[14, 14, 14, 14],
    )
    t_col_izquierda.setStyle(
        TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f7fafc")),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ])
    )

    archivos_logo = list(script_dir.glob("[Bb]elen.*")) + list(
        script_dir.glob("logo_derecha.*")
    )
    img_derecha_element = None
    if archivos_logo:
      try:
        ruta_img_logo = archivos_logo[0]
        pil_img_d = PILImage.open(ruta_img_logo).convert("RGBA")
        target_size = 120
        pil_img_d = pil_img_d.resize(
            (target_size, target_size), PILImage.Resampling.LANCZOS
        )

        mask = PILImage.new("L", (target_size, target_size), 0)
        draw = ImageDraw.Draw(mask)
        draw.ellipse((0, 0, target_size, target_size), fill=255)

        output = PILImage.new("RGBA", (target_size, target_size), (0, 0, 0, 0))
        output.paste(pil_img_d, (0, 0), mask=mask)

        img_d_io = io.BytesIO()
        output.save(img_d_io, format="PNG")
        img_d_io.seek(0)
        img_derecha_element = Image(img_d_io, width=45, height=45)
      except Exception:
        img_derecha_element = None

    textos_bloque = [
        Paragraph("MINISTERIO DE SALUD", title_main),
        Spacer(1, 1),
        Paragraph(
            "OFICINA GENERAL DE TECNOLOGIAS DE LA INFORMACION", title_sub
        ),
        Paragraph("UNIDAD DE ESTADISTICA RED DE SALUD HUAMANGA", title_sub),
        Spacer(1, 1),
        Paragraph(
            "Registro Diario de Atención y Otras Actividades de Salud", title_desc
        ),
    ]

    if img_derecha_element:
      centro_textos = Table(
          [[img_derecha_element, textos_bloque]], colWidths=[48, 219]
      )
      centro_textos.setStyle(
          TableStyle([
              ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
              ("ALIGN", (0, 0), (0, 0), "CENTER"),
              ("ALIGN", (1, 0), (1, 0), "CENTER"),
              ("LEFTPADDING", (0, 0), (-1, -1), 0),
              ("RIGHTPADDING", (0, 0), (-1, -1), 0),
              ("TOPPADDING", (0, 0), (-1, -1), 0),
              ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
          ])
      )
    else:
      centro_textos = textos_bloque

    t_turno_sub = Table(
        [
            [
                Paragraph("1", cell_center),
                Paragraph("TURNO", cell_center),
                Paragraph("M", cell_center),
                Paragraph("T", cell_center),
                Paragraph("N", cell_center),
            ],
            [
                Paragraph("", cell_center),
                Paragraph("", cell_center),
                Paragraph(m_check, cell_center),
                Paragraph(t_check, cell_center),
                Paragraph(n_check, cell_center),
            ],
        ],
        colWidths=[20, 45, 30, 30, 30],
        rowHeights=[10, 12],
    )
    t_turno_sub.setStyle(
        TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ])
    )

    header_top_data = [
        [t_col_izquierda, centro_textos, contenido_firma_celda],
        ["", "", t_turno_sub],
    ]

    t_header_top = Table(
        header_top_data, colWidths=[105, 272, 160], rowHeights=[45, 24]
    )
    t_header_top.setStyle(
        TableStyle([
            ("SPAN", (0, 0), (0, 1)),
            ("SPAN", (1, 0), (1, 1)),
            ("SPAN", (2, 0), (2, 0)),
            ("SPAN", (2, 1), (2, 1)),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("BACKGROUND", (2, 1), (2, 1), colors.HexColor("#ffffff")),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("LEFTPADDING", (0, 0), (-1, -1), 1),
            ("RIGHTPADDING", (0, 0), (-1, -1), 1),
        ])
    )

    elements.append(t_header_top)
    elements.append(Spacer(1, 4))

    mes_dict = {
        "Enero": "01",
        "Febrero": "02",
        "Marzo": "03",
        "Abril": "04",
        "Mayo": "05",
        "Junio": "06",
        "Julio": "07",
        "Agosto": "08",
        "Septiembre": "09",
        "Octubre": "10",
        "Noviembre": "11",
        "Diciembre": "12",
    }
    mes_num = mes_dict.get(mes, "09")
    anio_cort = anio[-2:] if len(anio) >= 2 else "26"

    header_bottom_data = [
        [
            Paragraph("2. AÑO", lbl_style),
            Paragraph("3. MES", lbl_style),
            Paragraph("4. NOMBRE DE ESTABLECIMIENTO DE SALUD (EESS)", lbl_style),
            Paragraph(
                "5. UNIDAD PRODUCTORA DE SERVICIOS (UPSS)", lbl_style
            ),
            Paragraph("6. NOMBRE DEL RESPONSABLE DE LA ATENCIÓN", lbl_style),
            Paragraph("DNI", lbl_style),
        ],
        [
            Paragraph(anio_cort, val_style),
            Paragraph(mes_num, val_style),
            Paragraph(centro_salud, val_style),
            Paragraph(ups_opcion, val_style),
            Paragraph(nombres_profesional, val_style),
            Paragraph(dni_profesional, val_style),
        ],
    ]

    t_header_bottom = Table(
        header_bottom_data, colWidths=[30, 30, 130, 120, 147, 80]
    )
    t_header_bottom.setStyle(
        TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#edf2f7")),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ])
    )

    elements.append(t_header_bottom)
    elements.append(Spacer(1, 6))

    headers = [
        Paragraph("N°", header_style),
        Paragraph("Día", header_style),
        Paragraph("DNI", header_style),
        Paragraph("Nombres y Apellidos del Paciente", header_style),
        Paragraph("Edad", header_style),
        Paragraph("Sexo", header_style),
        Paragraph("Antropometria", header_style),
        Paragraph("N/C/R", header_style),
        Paragraph("CIE / CPT", header_style),
        Paragraph("Diagnósticos", header_style),
        Paragraph("Lab", header_style),
        Paragraph("P/D/R", header_style),
    ]

    table_data = [headers]

    for p in st.session_state.lista_pacientes_medicina:
      sub_rows = []
      for cod, desc, lab, prd in zip(
          p["codigos"], p["descripciones"], p["labs"], p["prds"]
      ):
        sub_rows.append([
            Paragraph(cod, cell_center),
            Paragraph(desc, cell_style),
            Paragraph(lab, cell_center),
            Paragraph(prd, cell_center),
        ])

      sub_style = [
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("TOPPADDING", (0, 0), (-1, -1), 1),
          ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
          ("LEFTPADDING", (0, 0), (-1, -1), 1),
          ("RIGHTPADDING", (0, 0), (-1, -1), 1),
      ]

      for i in range(len(sub_rows) - 1):
        sub_style.append(
            ("LINEBELOW", (0, i), (-1, i), 0.25, colors.HexColor("#cbd5e0"))
        )

      t_sub_paciente = Table(sub_rows, colWidths=[45, 120, 20, 20])
      t_sub_paciente.setStyle(TableStyle(sub_style))

      table_data.append([
          Paragraph(str(p["nro"]), cell_center),
          Paragraph(p["dia_atencion"], cell_center),
          Paragraph(p["dni"], cell_center),
          Paragraph(p["nombres"], cell_style),
          Paragraph(str(p["edad"]), cell_edad),
          Paragraph(p["sexo"], cell_center),
          Paragraph(p["antropometria"], cell_center),
          Paragraph(p["condicion"], cell_center),
          t_sub_paciente,
          "",
          "",
          "",
      ])

    t = Table(
        table_data, colWidths=[18, 22, 50, 110, 25, 22, 70, 24, 45, 120, 20, 21]
    )

    t_style_commands = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2b6cb0")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        (
            "ROWBACKGROUNDS",
            (0, 1),
            (-1, -1),
            [colors.white, colors.HexColor("#f7fafc")],
        ),
    ]

    for row_idx in range(
        1, len(st.session_state.lista_pacientes_medicina) + 1
    ):
      t_style_commands.append(("SPAN", (8, row_idx), (11, row_idx)))
      t_style_commands.append(("LEFTPADDING", (8, row_idx), (11, row_idx), 0))
      t_style_commands.append(("RIGHTPADDING", (8, row_idx), (11, row_idx), 0))
      t_style_commands.append(("TOPPADDING", (8, row_idx), (11, row_idx), 0))
      t_style_commands.append(("BOTTOMPADDING", (8, row_idx), (11, row_idx), 0))

    t.setStyle(TableStyle(t_style_commands))
    elements.append(t)

    doc.build(elements)
    buffer.seek(0)

    st.download_button(
        label="📥 Descargar PDF A4 Vertical con Encabezado HIS Oficial",
        data=buffer,
        file_name="Registro_HIS_Medicina_Vertical.pdf",
        mime="application/pdf",
        type="primary",
        use_container_width=True,
    )

if st.button("🗑️ Limpiar Lista"):
  st.session_state.lista_pacientes_medicina = []
  st.rerun()

# --- PIE DE PÁGINA GENERAL DE LA APLICACIÓN ---
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("---")
st.markdown(
    "<p style='text-align: center; font-size: 12px; color: gray;'>Sistema HIS"
    " MINSA Local - Desarrollado por Carlos Sauñe Yupanqui ® 2026</p>",
    unsafe_allow_html=True,
)
