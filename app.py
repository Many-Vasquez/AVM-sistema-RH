from datetime import datetime, time
from io import BytesIO
import os
import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

# Configuración de la página
st.set_page_config(
    page_title="AVM - Sistema de Recursos Humanos y Comercial",
    page_icon="🛡️",
    layout="wide",
)

# Estilos corporativos en Negro y Dorado + visibilidad de botones
st.markdown(
    """
    <style>
        .stApp {
            background-color: #0e0e0e;
            color: #f3f3f3;
        }
        [data-testid="stSidebar"] {
            background-color: #161616;
            border-right: 1px solid #d4af37;
        }
        button[kind="header"] {
            color: #d4af37 !important;
            background-color: #1a1a1a !important;
            border: 1px solid #d4af37 !important;
        }
        [data-testid="collapsedControl"] {
            color: #d4af37 !important;
        }
        h1, h2, h3, h4, h5, h6, span, label {
            color: #f3f3f3 !important;
        }
        h1 {
            color: #d4af37 !important;
            border-bottom: 2px solid #d4af37;
            padding-bottom: 10px;
        }
        /* Botones generales estilizados y visibles */
        .stButton>button {
            background-color: #d4af37;
            color: #0e0e0e;
            font-weight: bold;
            border: none;
            border-radius: 6px;
            padding: 0.6rem 1.2rem;
            width: 100%;
        }
        .stButton>button:hover {
            background-color: #f3e5ab;
            color: #000000;
        }
        div.stForm {
            background-color: #1a1a1a;
            padding: 20px;
            border-radius: 8px;
            border: 1px solid #333333;
        }
        input, select, textarea {
            background-color: #222222 !important;
            color: #ffffff !important;
            border: 1px solid #444444 !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# --- BASE DE DATOS PERSISTENTE (CSV) ---
DB_FILE = "personal_avm.csv"
DB_ASISTENCIA = "asistencias_avm.csv"


def cargar_datos_empleados():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE).to_dict(orient="records")
    return []


def guardar_datos_empleados(lista_empleados):
    df = pd.DataFrame(lista_empleados)
    df.to_csv(DB_FILE, index=False)


def cargar_datos_asistencias():
    if os.path.exists(DB_ASISTENCIA):
        return pd.read_csv(DB_ASISTENCIA).to_dict(orient="records")
    return []


def guardar_datos_asistencias(lista_asistencias):
    df = pd.DataFrame(lista_asistencias)
    df.to_csv(DB_ASISTENCIA, index=False)


# Inicialización de estados sincronizados con archivos
if "empleados" not in st.session_state:
    st.session_state.empleados = cargar_datos_empleados()

if "asistencias" not in st.session_state:
    st.session_state.asistencias = cargar_datos_asistencias()

if "cotizacion_generada" not in st.session_state:
    st.session_state.cotizacion_generada = False


# --- MENÚ DE NAVEGACIÓN EN LA BARRA LATERAL ---
logo_path = (
    "Imagen1 (1).png" if os.path.exists("Imagen1 (1).png") else "logo.png"
)
if os.path.exists(logo_path):
    st.sidebar.image(logo_path, width=150)

st.sidebar.markdown("### 🧭 Menú de Navegación", unsafe_allow_html=True)

menu = st.sidebar.radio(
    "Seleccione el Módulo:",
    [
        "🏠 Inicio",
        "📊 Módulo Comercial (Cotizador)",
        "👥 Registro de Personal",
        "📥 Reporte de Personal (Excel)",
        "👆 Checador Biométrico de Huella",
        "📈 Reportes Métricos de Asistencia",
        "📄 Generación de Contratos",
    ],
)


# Función auxiliar para quitar bordes a tablas en Word
def remove_table_borders(table):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement("w:tblBorders")
    for border_name in [
        "top",
        "left",
        "bottom",
        "right",
        "insideH",
        "insideV",
    ]:
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "none")
        tblBorders.append(border)
    tblPr.append(tblBorders)


# --- 🏠 PÁGINA DE INICIO ---
if menu == "🏠 Inicio":
    col_c1, col_c2, col_c3 = st.columns([1, 2, 1])
    with col_c2:
        if os.path.exists(logo_path):
            st.image(logo_path, use_container_width=True)
        st.markdown(
            "<h1 style='text-align: center; color: #d4af37;'>AVM Grupo Integral de Seguridad Privada del Norte</h1>",
            unsafe_allow_html=True,
        )
        st.markdown(
            "<p style='text-align: center; color: #aaaaaa; font-size: 1.1rem;'>Sistema de Gestión Operativa, Recursos Humanos y Propuestas Comerciales</p>",
            unsafe_allow_html=True,
        )

# --- 📊 MÓDULO COMERCIAL (COTIZADOR) ---
elif menu == "📊 Módulo Comercial (Cotizador)":
    st.header("📊 Módulo Comercial - Generador de Cotizaciones")
    st.markdown("Complete los datos de la propuesta económica:")

    with st.form("form_cotizacion"):
        col1, col2 = st.columns(2)
        with col1:
            empresa_cliente = st.text_input(
                "Nombre de la Empresa Cliente", value=""
            )
            contacto_cliente = st.text_input(
                "Nombre del Contacto / Comprador", value=""
            )
            fecha_cot = st.text_input("Fecha de Emisión", value="")
        with col2:
            cantidad_guardias = st.number_input(
                "Cantidad de Guardias",
                min_value=0,
                max_value=50,
                value=0,
                step=1,
            )
            precio_unitario = st.number_input(
                "Precio Unitario Mensual por Guardia ($)",
                min_value=0.0,
                value=0.0,
                step=100.0,
                format="%.2f",
            )

        submitted_cot = st.form_submit_button("⚙️ Generar Propuesta Económica")

        if submitted_cot:
            if empresa_cliente and cantidad_guardias > 0 and precio_unitario > 0:
                st.session_state.cotizacion_generada = True
                st.session_state.empresa_cliente = empresa_cliente
                st.session_state.contacto_cliente = contacto_cliente
                st.session_state.fecha_cot = fecha_cot
                st.session_state.cantidad_guardias = cantidad_guardias
                st.session_state.precio_unitario = precio_unitario
                st.success("¡Datos de cotización cargados con éxito!")
            else:
                st.error(
                    "Por favor complete el nombre del cliente, cantidad de guardias y precio unitario."
                )

    if st.session_state.get("cotizacion_generada", False):
        subtotal = (
            st.session_state.cantidad_guardias
            * st.session_state.precio_unitario
        )
        iva = subtotal * 0.16
        total = subtotal + iva

        doc_cot = Document()

        for section in doc_cot.sections:
            section.top_margin = Inches(0.8)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(1.0)
            section.right_margin = Inches(1.0)

        COLOR_DORADO = RGBColor(197, 155, 39)
        COLOR_NEGRO_SUAVE = RGBColor(20, 20, 20)
        COLOR_GRIS_TEXTO = RGBColor(80, 80, 80)

        header_table = doc_cot.add_table(rows=1, cols=2)
        header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_table_borders(header_table)

        cell_logo = header_table.cell(0, 0)
        cell_logo.width = Inches(1.2)
        p_logo = cell_logo.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if os.path.exists(logo_path):
            p_logo.add_run().add_picture(logo_path, width=Inches(1.0))

        cell_text = header_table.cell(0, 1)
        cell_text.width = Inches(5.3)

        p_emp = cell_text.paragraphs[0]
        p_emp.paragraph_format.space_after = Pt(1)
        run_emp_1 = p_emp.add_run("AVM")
        run_emp_1.bold = True
        run_emp_1.font.size = Pt(13)
        run_emp_1.font.color.rgb = COLOR_NEGRO_SUAVE

        run_emp_2 = p_emp.add_run(
            " GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE, S.A. DE C.V."
        )
        run_emp_2.bold = True
        run_emp_2.font.size = Pt(11)
        run_emp_2.font.color.rgb = COLOR_DORADO

        p_dir = cell_text.add_paragraph()
        p_dir.paragraph_format.space_after = Pt(0)
        run_dir = p_dir.add_run(
            "SANTA BÁRBARA NÚMERO 141, COLONIA VALLE DE SANTA ISABEL, C.P. 67256,\nCIUDAD BENITO JUÁREZ, NUEVO LEÓN"
        )
        run_dir.font.size = Pt(7.5)
        run_dir.font.color.rgb = COLOR_GRIS_TEXTO

        p_line = doc_cot.add_paragraph()
        p_line.paragraph_format.space_before = Pt(4)
        p_line.paragraph_format.space_after = Pt(6)
        r_line = p_line.add_run(
            "_________________________________________________________________________________"
        )
        r_line.font.size = Pt(8)
        r_line.font.color.rgb = COLOR_DORADO

        p_prop = doc_cot.add_paragraph()
        p_prop.paragraph_format.space_after = Pt(4)
        run_prop = p_prop.add_run("PROPUESTA ECONÓMICA DE SERVICIOS")
        run_prop.bold = True
        run_prop.font.size = Pt(11)
        run_prop.font.color.rgb = COLOR_DORADO

        p_datos = doc_cot.add_paragraph()
        p_datos.paragraph_format.space_after = Pt(6)
        p_datos.add_run(f"FECHA DE EMISIÓN:  {st.session_state.fecha_cot}\n")
        p_datos.add_run(
            f"CLIENTE:                  {st.session_state.empresa_cliente}\n"
        )
        p_datos.add_run(
            f"ATENCIÓN:               {st.session_state.contacto_cliente}\n"
        )
        for run in p_datos.runs:
            run.font.size = Pt(9)
            run.bold = True
            run.font.color.rgb = COLOR_NEGRO_SUAVE

        h2_1 = doc_cot.add_heading(level=2)
        h2_1.paragraph_format.space_before = Pt(2)
        h2_1.paragraph_format.space_after = Pt(2)
        r_h2_1 = h2_1.add_run("ANÁLISIS DE SITUACIÓN:")
        r_h2_1.font.size = Pt(10)
        r_h2_1.font.color.rgb = COLOR_DORADO

        p_analisis = doc_cot.add_paragraph(
            "Tras evaluar las necesidades de seguridad de su instalación, nuestra firma propone un esquema de Seguridad Proactiva. A diferencia de la vigilancia convencional, nuestro servicio se basa en la disuasión avanzada y la capacidad de respuesta inmediata bajo los más altos estándares de cumplimiento legal."
        )
        p_analisis.paragraph_format.space_after = Pt(6)
        p_analisis.runs[0].font.size = Pt(9)
        p_analisis.runs[0].font.color.rgb = COLOR_NEGRO_SUAVE

        h2_2 = doc_cot.add_heading(level=2)
        h2_2.paragraph_format.space_before = Pt(2)
        h2_2.paragraph_format.space_after = Pt(2)
        r_h2_2 = h2_2.add_run("DETALLE DE COTIZACIÓN:")
        r_h2_2.font.size = Pt(10)
        r_h2_2.font.color.rgb = COLOR_DORADO

        table = doc_cot.add_table(rows=2, cols=5)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        headers = [
            "CANT.",
            "CATEGORÍA",
            "DESCRIPCIÓN DEL SERVICIO",
            "PRECIO UNITARIO",
            "TOTAL MENSUAL",
        ]
        hdr_cells = table.rows[0].cells
        for i, header_text in enumerate(headers):
            hdr_cells[i].text = header_text
            for paragraph in hdr_cells[i].paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.size = Pt(8)
                    run.font.color.rgb = RGBColor(255, 255, 255)
            shading = OxmlElement("w:shd")
            shading.set(qn("w:val"), "clear")
            shading.set(qn("w:color"), "auto")
            shading.set(qn("w:fill"), "C59B27")
            hdr_cells[i]._tc.get_or_add_tcPr().append(shading)

        row_cells = table.rows[1].cells
        row_cells[0].text = str(st.session_state.cantidad_guardias)
        row_cells[1].text = (
            "Guardias Intramuros/Extramuros- Control de Accesos"
        )
        row_cells[2].text = (
            "Control estricto de acceso peatonal y vehicular "
            "(empleados, contratistas, proveedores y transporte pesado). Turno de 12 horas."
        )
        row_cells[3].text = f"${st.session_state.precio_unitario:,.2f}"
        row_cells[4].text = f"${subtotal:,.2f}"

        for i, cell in enumerate(row_cells):
            for paragraph in cell.paragraphs:
                if i in [0, 3, 4]:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for run in paragraph.runs:
                    run.font.size = Pt(8)
                    run.font.color.rgb = COLOR_NEGRO_SUAVE

        p_totales = doc_cot.add_paragraph()
        p_totales.paragraph_format.space_before = Pt(4)
        p_totales.paragraph_format.space_after = Pt(6)
        p_totales.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_totales.add_run(f"Subtotal: ${subtotal:,.2f}\n")
        p_totales.add_run(f"IVA (16%): ${iva:,.2f}\n")
        r_tot = p_totales.add_run(f"TOTAL MENSUAL: ${total:,.2f}")
        r_tot.bold = True
        r_tot.font.size = Pt(10)
        r_tot.font.color.rgb = COLOR_DORADO

        for run in p_totales.runs:
            if run != r_tot:
                run.font.size = Pt(8.5)
                run.font.color.rgb = COLOR_NEGRO_SUAVE

        h2_3 = doc_cot.add_heading(level=2)
        h2_3.paragraph_format.space_before = Pt(2)
        h2_3.paragraph_format.space_after = Pt(2)
        r_h2_3 = h2_3.add_run("TÉRMINOS Y CONDICIONES COMERCIALES:")
        r_h2_3.font.size = Pt(10)
        r_h2_3.font.color.rgb = COLOR_DORADO

        terminos = [
            (
                "1. Responsabilidad Civil y Patronal:",
                "Nuestra firma asume la totalidad de las obligaciones derivadas de las leyes laborales, de seguridad social (IMSS, INFONAVIT) y fiscales vigentes.",
            ),
            (
                "2. Garantía de Continuidad:",
                "Nos comprometemos a mantener la cobertura del servicio al 100%. Sustitución en menos de 90 minutos por personal de retén.",
            ),
            (
                "3. Confidencialidad Rigurosa:",
                "Todo el personal asignado cuenta con estrictos contratos de confidencialidad para proteger las operaciones del cliente.",
            ),
            (
                "4. Vigencia de la Propuesta:",
                "La presente cotización tiene una validez de 15 días naturales a partir de su emisión.",
            ),
            (
                "5. Condiciones de Pago:",
                "Facturación mensual liquidable dentro de los primeros 5 días naturales de cada mes.",
            ),
        ]

        for titulo, desc in terminos:
            p_term = doc_cot.add_paragraph()
            p_term.paragraph_format.space_after = Pt(1)
            r_t = p_term.add_run(titulo + " ")
            r_t.bold = True
            r_t.font.size = Pt(8)
            r_t.font.color.rgb = COLOR_DORADO

            r_d = p_term.add_run(desc)
            r_d.font.size = Pt(8)
            r_d.font.color.rgb = COLOR_NEGRO_SUAVE

        p_pie = doc_cot.add_paragraph()
        p_pie.paragraph_format.space_before = Pt(6)
        p_pie.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_pie = p_pie.add_run(
            '"Nuestra estructura operativa garantiza que el error humano se reduzca al mínimo mediante la supervisión cruzada y el respaldo tecnológico en tiempo real."'
        )
        r_pie.italic = True
        r_pie.font.size = Pt(8)
        r_pie.font.color.rgb = COLOR_DORADO

        buffer_cot = BytesIO()
        doc_cot.save(buffer_cot)
        buffer_cot.seek(0)

        st.success("¡Propuesta económica generada con éxito!")
        st.download_button(
            label="📥 Descargar Propuesta Económica en Word",
            data=buffer_cot,
            file_name=f"Cotizacion_AVM_{st.session_state.empresa_cliente.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

# --- 👥 REGISTRO DE PERSONAL ---
elif menu == "👥 Registro de Personal":
    st.header("📝 Registro de Nuevo Elemento / Guardia")

    with st.form("form_empleado"):
        col1, col2 = st.columns(2)

        with col1:
            nombre = st.text_input("Nombre Completo del Trabajador")
            nacionalidad = st.text_input("Nacionalidad", value="Mexicana")
            sexo = st.selectbox("Sexo", ["Masculino", "Femenino"])
            fecha_nacimiento = st.text_input(
                "Fecha de Nacimiento (ej. 15/05/1995)"
            )
            estado_civil = st.selectbox(
                "Estado Civil", ["Soltero/a", "Casado/a", "Viudo/a"]
            )

        with col2:
            curp = st.text_input("CURP")
            rfc = st.text_input("RFC")
            nss = st.text_input("NSS (Número de Seguridad Social - 11 dígitos)")
            domicilio = st.text_area("Domicilio Completo")
            puesto = st.text_input("Puesto", value="GUARDIA DE SEGURIDAD")
            salario_semanal = st.text_input(
                "Salario Semanal", value="$2,103.85"
            )

        submitted = st.form_submit_button("💾 Guardar Elemento en Base de Datos")

        if submitted:
            if nombre and curp and nss:
                nuevo_emp = {
                    "Nombre": nombre,
                    "Nacionalidad": nacionalidad,
                    "Sexo": sexo,
                    "Fecha de Nacimiento": fecha_nacimiento,
                    "Estado Civil": estado_civil,
                    "CURP": curp,
                    "RFC": rfc,
                    "NSS": nss,
                    "Domicilio": domicilio,
                    "Puesto": puesto,
                    "Salario Semanal": salario_semanal,
                }
                st.session_state.empleados.append(nuevo_emp)
                guardar_datos_empleados(st.session_state.empleados)
                st.success(
                    f"¡Guardia {nombre} registrado y guardado en la base de datos central!"
                )
            else:
                st.error("Por favor complete al menos Nombre, CURP y NSS.")

    # Recargar datos frescos del archivo CSV para ver cambios en tiempo real
    st.session_state.empleados = cargar_datos_empleados()
    if len(st.session_state.empleados) > 0:
        st.subheader("📋 Plantilla de Personal Registrado (Sincronizado)")
        df_personal = pd.DataFrame(st.session_state.empleados)
        st.dataframe(df_personal, use_container_width=True)

# --- 📥 REPORTE DE PERSONAL (CSV / EXCEL) ---
elif menu == "📥 Reporte de Personal (Excel)":
    st.header("📥 Módulo de Reportes de Personal")
    st.markdown(
        "Genere y descargue el padrón completo del personal activo para auditoría o administración."
    )

    st.session_state.empleados = cargar_datos_empleados()
    if not st.session_state.empleados:
        st.info(
            "ℹ️ No hay registros de personal en este momento. Registre personal en la sección previa."
        )
    else:
        df_excel = pd.DataFrame(st.session_state.empleados)
        st.dataframe(df_excel, use_container_width=True)

        csv_personal = df_excel.to_csv(index=False).encode("utf-8")

        st.download_button(
            label="📊 Descargar Listado de Personal (Compatible con Excel)",
            data=csv_personal,
            file_name=f"Padron_Personal_AVM_{datetime.now().strftime('%Y-%m-%d')}.csv",
            mime="text/csv",
        )

# --- 👆 CHECADOR BIOMÉTRICO DE HUELLA ---
elif menu == "👆 Checador Biométrico de Huella":
    st.header("👆 Terminal Biométrica - Control de Asistencia y Turnos")
    st.markdown(
        "Módulo independiente para registro de huella y cambios de turno en planta."
    )

    st.session_state.empleados = cargar_datos_empleados()
    if not st.session_state.empleados:
        st.warning(
            "⚠️ No hay elementos registrados. Registre personal primero en el módulo de 'Registro de Personal'."
        )
    else:
        col_b1, col_b2 = st.columns(2)

        with col_b1:
            nombres_empleados = [
                e["Nombre"] for e in st.session_state.empleados
            ]
            emp_checador = st.selectbox(
                "Seleccionar Elemento / Guardia", nombres_empleados
            )
            datos_emp = next(
                e
                for e in st.session_state.empleados
                if e["Nombre"] == emp_checador
            )

            st.info(
                f"**NSS:** {datos_emp['NSS']} \n\n **Puesto:** {datos_emp['Puesto']}"
            )

            tipo_movimiento = st.radio(
                "Tipo de Registro", ["Entrada de Turno", "Salida de Turno"]
            )

        with col_b2:
            st.markdown("### 🖐️ Validación Biométrica")
            st.markdown(
                "Coloque el dedo en el lector USB o presione para simular fichaje:"
            )

            if st.button(
                "🔴 ESCANEAR HUELLA DIGITAL (Simulador)",
                use_container_width=True,
            ):
                ahora = datetime.now()
                fecha_str = ahora.strftime("%Y-%m-%d")
                hora_str = ahora.strftime("%H:%M:%S")

                hora_limite = time(8, 0, 0)
                hora_actual = ahora.time()

                if tipo_movimiento == "Entrada de Turno":
                    if hora_actual <= hora_limite:
                        puntualidad = "A TIEMPO"
                        cumple_puntualidad = "SÍ"
                    else:
                        puntualidad = "RETARDO"
                        cumple_puntualidad = "NO"
                else:
                    puntualidad = "N/A (Salida)"
                    cumple_puntualidad = "N/A"

                if (
                    cumple_puntualidad == "SÍ"
                    or tipo_movimiento == "Salida de Turno"
                ):
                    bono_otorgado = "SÍ ($900.00)"
                else:
                    bono_otorgado = "NO (Castigado por retardo)"

                nuevo_registro = {
                    "Fecha": fecha_str,
                    "Nombre": datos_emp["Nombre"],
                    "NSS": datos_emp["NSS"],
                    "Movimiento": tipo_movimiento,
                    "Hora": hora_str,
                    "Estado": puntualidad,
                    "Bono Asistencia/Puntualidad": bono_otorgado,
                }

                st.session_state.asistencias.append(nuevo_registro)
                guardar_datos_asistencias(st.session_state.asistencias)
                st.success(
                    f"✅ ¡{tipo_movimiento.upper()} registrada para {datos_emp['Nombre']} a las {hora_str}!"
                )

    st.session_state.asistencias = cargar_datos_asistencias()
    if len(st.session_state.asistencias) > 0:
        st.subheader("⚡ Últimos Registros Biométricos (Sincronizados)")
        df_asist = pd.DataFrame(st.session_state.asistencias)
        st.dataframe(df_asist, use_container_width=True)

# --- 📈 REPORTES MÉTRICOS DE ASISTENCIA ---
elif menu == "📈 Reportes Métricos de Asistencia":
    st.header("📈 Reportes Métricos y Auditoría de Asistencia")
    st.markdown(
        "Control gerencial de incidencias, retardos y bonos semanales en tiempo real."
    )

    st.session_state.asistencias = cargar_datos_asistencias()
    if not st.session_state.asistencias:
        st.info("ℹ️ Aún no hay registros en el checador biométrico.")
    else:
        df_reportes = pd.DataFrame(st.session_state.asistencias)

        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            st.metric(label="Total Registros Biométricos", value=len(df_reportes))
        with col_m2:
            retardos = len(df_reportes[df_reportes["Estado"] == "RETARDO"])
            st.metric(
                label="Incidencias / Retardos",
                value=retardos,
                delta=f"-{retardos}" if retardos > 0 else "0",
                delta_color="inverse",
            )
        with col_m3:
            bonos_ok = len(
                df_reportes[
                    df_reportes["Bono Asistencia/Puntualidad"].str.contains(
                        "SÍ"
                    )
                ]
            )
            st.metric(label="Bonos Ganados", value=bonos_ok)

        st.subheader("📋 Tabla de Auditoría Operativa")
        st.dataframe(df_reportes, use_container_width=True)

        csv_data = df_reportes.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Exportar Reporte de Asistencia a CSV (Excel)",
            data=csv_data,
            file_name=f"Reporte_Asistencia_AVM_{datetime.now().strftime('%Y-%m-%d')}.csv",
            mime="text/csv",
        )

# --- 📄 GENERACIÓN DE CONTRATOS ---
elif menu == "📄 Generación de Contratos":
    st.header("📄 Generador de Contratos Laborales en Word")

    st.session_state.empleados = cargar_datos_empleados()
    if not st.session_state.empleados:
        st.warning(
            "⚠️ Primero registre personal en la sección 'Registro de Personal'."
        )
    else:
        tipo_contrato = st.radio(
            "Seleccione Tipo de Contrato a Generar",
            ["Sujeto a Prueba (30 Días)", "Tiempo Indeterminado"],
        )
        nombres_empleados = [e["Nombre"] for e in st.session_state.empleados]
        emp_seleccionado = st.selectbox(
            "Seleccionar Trabajador", nombres_empleados
        )
        datos = next(
            e
            for e in st.session_state.empleados
            if e["Nombre"] == emp_seleccionado
        )

        if st.button("📥 Generar y Descargar Contrato en Word"):
            doc = Document()
            titulo_contrato = (
                "CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO INDETERMINADO, SUJETO A UN PERIODO DE PRUEBA"
                if "Prueba" in tipo_contrato
                else "CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO INDETERMINADO"
            )
            doc.add_heading(titulo_contrato, level=1)

            texto_cuerpo = f"""CONTRATO INDIVIDUAL DE TRABAJO QUE CELEBRAN, POR UNA PARTE, AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE, S.A. DE C.V., Y POR LA OTRA, {datos['Nombre']}, AL TENOR DE LAS SIGUIENTES CLAUSULAS Y DECLARACIONES:

I. Declara el PATRÓN ser una empresa legalmente constituida con domicilio en Calle Santa Bárbara número 141, Colonia Valle de Santa Isabel, C.P. 67256, Ciudad Benito Juárez, Nuevo León.

II. Declara el TRABAJADOR llamarse {datos['Nombre']}, de nacionalidad {datos['Nacionalidad']}, sexo {datos['Sexo']}, estado civil {datos['Estado Civil']}, con CURP {datos['CURP']}, RFC {datos['RFC']}, NSS {datos['NSS']} y domicilio en {datos['Domicilio']}.

CLÁUSULAS:
PRIMERA. El TRABAJADOR prestará sus servicios desempeñando el puesto de {datos['Puesto']}.
SEGUNDA. El salario semanal será de {datos['Salario Semanal']}, cubriéndose los viernes de cada semana. Se otorgan bonos de asistencia y puntualidad condicionados al registro biométrico puntual.
TERCERA. Las demás condiciones se rigen por la Ley Federal del Trabajo y el Reglamento Interior de Trabajo.

Se firma por duplicado en Ciudad Benito Juárez, Nuevo León.

EL PATRÓN
AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE S.A. DE C.V.
C. ABNER VELAZQUEZ MORALES

__________________________________

EL TRABAJADOR
{datos['Nombre']}

__________________________________
"""
            for parrafo in texto_cuerpo.split("\n\n"):
                p_limpio = parrafo.strip()
                if p_limpio:
                    doc.add_paragraph(p_limpio)

            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.download_button(
                label="📥 Descargar Documento Word Listo",
                data=buffer,
                file_name=f"Contrato_{tipo_contrato.split()[0]}_{datos['Nombre'].replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
