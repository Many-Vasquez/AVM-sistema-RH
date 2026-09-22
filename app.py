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

# Estilos corporativos en Negro y Dorado
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
        .stButton>button {
            background-color: #d4af37;
            color: #0e0e0e;
            font-weight: bold;
            border: none;
            border-radius: 4px;
            padding: 0.5rem 1rem;
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

# Encabezado con el Logotipo
col_logo, col_titulo = st.columns([1, 4])

with col_logo:
    if os.path.exists("Imagen1 (1).png"):
        st.image("Imagen1 (1).png", width=130)
    elif os.path.exists("logo.png"):
        st.image("logo.png", width=130)
    else:
        st.markdown("🛡️ **[Sube tu logo al repositorio]**")

with col_titulo:
    st.title("AVM Grupo Integral de Seguridad Privada del Norte")
    st.subheader(
        "Sistema de Gestión de Recursos Humanos y Propuestas Comerciales"
    )

# --- INICIALIZACIÓN DE ESTADOS ---
if "empleados" not in st.session_state:
    st.session_state.empleados = []

if "asistencias" not in st.session_state:
    st.session_state.asistencias = []

if "cotizacion_generada" not in st.session_state:
    st.session_state.cotizacion_generada = False

if "vista_actual" not in st.session_state:
    st.session_state.vista_actual = "Sistema de Recursos Humanos"

# --- MENÚ LATERAL ---
st.sidebar.markdown("### 🧭 Administración", unsafe_allow_html=True)

if st.sidebar.button("📊 Ir a Módulo Comercial (Cotizador)"):
    st.session_state.vista_actual = "Generador de Cotizaciones"

menu_hr = st.sidebar.selectbox(
    "Módulo de Recursos Humanos",
    [
        "Registro de Personal",
        "Checador Biométrico de Huella",
        "Reportes Métricos de Asistencia",
        "Generar Contrato Sujeto a Prueba",
        "Generar Contrato Tiempo Indeterminado",
    ],
)

if menu_hr != st.session_state.get("menu_hr_prev", ""):
    st.session_state.vista_actual = "RRHH"
    st.session_state.menu_hr_prev = menu_hr


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


# --- LÓGICA DE VISTAS ---
if st.session_state.vista_actual == "Generador de Cotizaciones":
    st.header("📊 Módulo Comercial - Generador de Cotizaciones para Clientes")

    with st.form("form_cotizacion"):
        col1, col2 = st.columns(2)
        with col1:
            empresa_cliente = st.text_input(
                "Nombre de la Empresa Cliente", value="CROMIN DE MEXICO"
            )
            contacto_cliente = st.text_input(
                "Nombre del Contacto / Comprador", value="LIC CRISTINA LIERA"
            )
            fecha_cot = st.text_input("Fecha de Emisión", value="11/09/2026")
        with col2:
            cantidad_guardias = st.number_input(
                "Cantidad de Guardias", min_value=1, max_value=50, value=2
            )
            precio_unitario = st.number_input(
                "Precio Unitario Mensual por Guardia ($)",
                min_value=0.0,
                value=21551.0,
                step=100.0,
            )

        submitted_cot = st.form_submit_button("⚙️ Generar Propuesta Económica")

        if submitted_cot:
            st.session_state.cotizacion_generada = True
            st.session_state.empresa_cliente = empresa_cliente
            st.session_state.contacto_cliente = contacto_cliente
            st.session_state.fecha_cot = fecha_cot
            st.session_state.cantidad_guardias = cantidad_guardias
            st.session_state.precio_unitario = precio_unitario

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
        logo_path = (
            "Imagen1 (1).png"
            if os.path.exists("Imagen1 (1).png")
            else "logo.png"
        )
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
            "Guardias Intramuro/Extramuros- Control de Accesos"
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

else:
    # --- MÓDULO DE RECURSOS HUMANOS Y ASISTENCIA ---
    if menu_hr == "Registro de Personal":
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
                nss = st.text_input(
                    "NSS (Número de Seguridad Social - 11 dígitos)"
                )
                domicilio = st.text_area("Domicilio Completo")
                puesto = st.text_input("Puesto", value="GUARDIA DE SEGURIDAD")
                salario_semanal = st.text_input(
                    "Salario Semanal (ej. $2,103.85)", value="$2,103.85"
                )

            submitted = st.form_submit_button("Guardar Empleado")

            if submitted:
                if nombre and curp:
                    nuevo_emp = {
                        "nombre": nombre,
                        "nacionalidad": nacionalidad,
                        "sexo": sexo,
                        "fecha_nacimiento": fecha_nacimiento,
                        "estado_civil": estado_civil,
                        "curp": curp,
                        "rfc": rfc,
                        "nss": nss,
                        "domicilio": domicilio,
                        "puesto": puesto,
                        "salario_semanal": salario_semanal,
                    }
                    st.session_state.empleados.append(nuevo_emp)
                    st.success(
                        f"¡Guardia {nombre} registrado correctamente con NSS {nss}!"
                    )
                else:
                    st.error(
                        "Por favor ingresa al menos el Nombre y la CURP."
                    )

        if len(st.session_state.empleados) > 0:
            st.subheader("📋 Personal Registrado Recientemente")
            df = pd.DataFrame(st.session_state.empleados)
            st.dataframe(df)

    elif menu_hr == "Checador Biométrico de Huella":
        st.header("👆 Terminal Biométrica - Control de Asistencia y Turnos")
        st.markdown(
            "Simulación de lectura de huella digital o registro rápido por número de seguridad social (NSS) para cambios de turno en planta."
        )

        if not st.session_state.empleados:
            st.warning(
                "⚠️ No hay elementos registrados. Registra personal primero en 'Registro de Personal'."
            )
        else:
            col_b1, col_b2 = st.columns(2)

            with col_b1:
                nombres_empleados = [
                    e["nombre"] for e in st.session_state.empleados
                ]
                emp_checador = st.selectbox(
                    "Seleccionar Elemento / Guardia", nombres_empleados
                )
                datos_emp = next(
                    e
                    for e in st.session_state.empleados
                    if e["nombre"] == emp_checador
                )

                st.info(
                    f"**NSS:** {datos_emp['nss']} \n\n **Puesto:** {datos_emp['puesto']}"
                )

                tipo_movimiento = st.radio(
                    "Tipo de Registro", ["Entrada de Turno", "Salida de Turno"]
                )

            with col_b2:
                st.markdown("### 🖐️ Panel de Validación Biométrica")
                st.markdown(
                    "Coloque el dedo en el lector biométrico o pulse el botón para registrar huella:"
                )

                # Simulación visual del lector de huella
                if st.button(
                    "🔴 ESCANEAR HUELLA DIGITAL (Simulador USB)",
                    use_container_width=True,
                ):
                    ahora = datetime.now()
                    fecha_str = ahora.strftime("%Y-%m-%d")
                    hora_str = ahora.strftime("%H:%M:%S")

                    # Regla de puntualidad: Entrada límite a las 08:00:00 AM para turno matutino
                    hora_limite = time(8, 0, 0)
                    hora_actual = ahora.time()

                    if tipo_movimiento == "Entrada de Turno":
                        if hora_actual <= hora_limite:
                            puntualidad = "A TIEMPO"
                            cumple_puntualidad = "SÍ"
                        else:
                            puntualidad = "RETARDO"
                            cumple_puntualidad = "NO"
                        asistencia_valida = "SÍ"
                    else:
                        puntualidad = "N/A (Salida)"
                        cumple_puntualidad = "N/A"
                        asistencia_valida = "SÍ"

                    # Bono de asistencia y puntualidad condicionado
                    if (
                        cumple_puntualidad == "SÍ"
                        or tipo_movimiento == "Salida de Turno"
                    ):
                        bono_otorgado = "SÍ ($900.00)"
                    else:
                        bono_otorgado = (
                            "NO (Castigado por retardo/inasistencia)"
                        )

                    nuevo_registro = {
                        "Fecha": fecha_str,
                        "Nombre": datos_emp["nombre"],
                        "NSS": datos_emp["nss"],
                        "Movimiento": tipo_movimiento,
                        "Hora": hora_str,
                        "Estado": puntualidad,
                        "Bono Asistencia/Puntualidad": bono_otorgado,
                    }

                    st.session_state.asistencias.append(nuevo_registro)
                    st.success(
                        f"✅ ¡{tipo_movimiento.upper()} registrada con éxito para {datos_emp['nombre']} a las {hora_str}!"
                    )

        if len(st.session_state.asistencias) > 0:
            st.subheader("⚡ Últimos Registros Biométricos en Tiempo Real")
            df_asist = pd.DataFrame(st.session_state.asistencias)
            st.dataframe(df_asist, use_container_width=True)

    elif menu_hr == "Reportes Métricos de Asistencia":
        st.header("📈 Reportes Métricos y Auditoría de Asistencia")
        st.markdown(
            "Control gerencial de incidencias, retardos y asignación de bonos semanales."
        )

        if not st.session_state.asistencias:
            st.info(
                "ℹ️ Aún no hay registros en el checador biométrico. Realiza registros en la sección anterior."
            )
        else:
            df_reportes = pd.DataFrame(st.session_state.asistencias)

            # Métricas rápidas en columnas
            col_m1, col_m2, col_m3 = st.columns(3)
            with col_m1:
                st.metric(
                    label="Total Registros Biométricos",
                    value=len(df_reportes),
                )
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

            st.subheader("📋 Tabla Completa de Auditoría Operativa")
            st.dataframe(df_reportes, use_container_width=True)

            # Botón de exportación a CSV para Excel
            csv_data = df_reportes.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Exportar Reporte Métrico a CSV (Compatible con Excel)",
                data=csv_data,
                file_name=f"Reporte_Asistencia_AVM_{datetime.now().strftime('%Y-%m-%d')}.csv",
                mime="text/csv",
            )

    elif menu_hr == "Generar Contrato Sujeto a Prueba":
        st.header("📄 Generador de Contrato - Sujeto a Prueba (30 Días)")

        if not st.session_state.empleados:
            st.warning(
                "⚠️ Primero registra un empleado en la sección 'Registro de Personal'."
            )
        else:
            nombres_empleados = [
                e["nombre"] for e in st.session_state.empleados
            ]
            emp_seleccionado = st.selectbox(
                "Selecciona al Trabajador", nombres_empleados
            )
            datos = next(
                e
                for e in st.session_state.empleados
                if e["nombre"] == emp_seleccionado
            )

            if st.button("📥 Descargar Contrato Sujeto a Prueba (Word)"):
                doc = Document()
                doc.add_heading(
                    "CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO INDETERMINADO, SUJETO A UN PERIODO DE PRUEBA",
                    level=1,
                )

                texto_prueba = f"""CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO INDETERMINADO, SUJETO A UN PERIODO DE PRUEBA, QUE CELEBRAN, POR UNA PARTE, AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE, SOCIEDAD ANONIMA DE CAPITAL VARIABLE, REPRESENTADA EN ESTE ACTO POR EL C. ABNER VELAZQUEZ MORALES (EN LO SUCESIVO, EL "PATRÓN"), Y POR LA OTRA PARTE, POR SU PROPIO DERECHO, {datos['nombre']} (EN LO SUCESIVO, EL “TRABAJADOR”), DE CONFORMIDAD CON LOS ARTÍCULOS 20, 21, 24, 25, 35, 39-A, 39-B, 132, 134 Y DEMÁS RELATIVOS Y APLICABLES DE LA LEY FEDERAL DEL TRABAJO, AL TENOR DE LAS SIGUIENTES DECLARACIONES Y CLÁUSULAS:

D E C L A R A C I O N E S:

I. Declara el PATRÓN:
a) Ser una persona moral, debidamente constituida conforme a las leyes de la República Mexicana, según consta en la escritura pública número 6,948, pasada ante la fe del Notario Público número 127, con domicilio ubicado en Calle Santa Bárbara número 141, C. Asturias, Colonia Valle de Santa Isabel, C.P. 67256, Ciudad Benito Juárez, Nuevo León, Registro Federal de Contribuyentes AGI260413CK4 y tener como objeto social, entre otros, la prestación de servicios de seguridad privada, consistentes en la vigilancia, protección y resguardo de bienes muebles e inmuebles, así como de establecimientos comerciales, industriales, habitacionales y de servicios, mediante la utilización de recursos humanos debidamente capacitados.
b) Que, para dar cumplimiento a su objeto social, requiere de personal capacitado y con experiencia para ocupar el puesto de {datos['puesto']} para que realice las actividades consistentes, de manera enunciativa mas no limitativa, en: Vigilancia, protección y resguardo de bienes muebles e inmuebles, seguridad intramuros incluyendo el control de accesos y salidas, registro de personas y vehículos, realización de rondines y supervisión interna de instalaciones privadas.

II. Declara el TRABAJADOR:
a) Ser una persona física, de nacionalidad {datos['nacionalidad']}, de sexo {datos['sexo']}, con fecha de nacimiento el {datos['fecha_nacimiento']}, estado civil {datos['estado_civil']}, Clave Única de Registro de Población {datos['curp']}, Registro Federal de Contribuyentes {datos['rfc']} y Número de Seguridad Social (NSS) {datos['nss']}, con domicilio en {datos['domicilio']}.
b) Que cuenta con los conocimientos, habilidades y experiencia necesarios para prestar al PATRÓN los servicios del puesto encomendado.
c) Que está de acuerdo en prestar los servicios descritos en el presente contrato, sujeto a un periodo de prueba de 30 (treinta) días.

III. Declaran ambas partes:
a) Que cuentan con las facultades suficientes para la celebración del presente contrato y obligarse a los términos de este, reconociéndose mutuamente la personalidad con la que comparecen.

C L Á U S U L A S:

PRIMERA. El presente contrato se celebra por TIEMPO INDETERMINADO, quedando sujeto "EL TRABAJADOR" a un PERIODO DE PRUEBA DE 1 MES (30 DÍAS) contados a partir de la fecha de firma del presente contrato, con fundamento en el párrafo segundo del Artículo 39-A de la Ley Federal del Trabajo. Durante dicho periodo de prueba, "EL PATRÓN" evaluará si "EL TRABAJADOR" cumple con los requisitos y aptitudes necesarios para el puesto.

SEGUNDA. Se hace constar que el PATRÓN celebra el presente contrato fundado en las declaraciones del TRABAJADOR. Al término del periodo de prueba, de no acreditar el TRABAJADOR que satisface los requisitos necesarios, se dará por terminada la relación de trabajo sin responsabilidad para el PATRÓN.

TERCERA. El TRABAJADOR prestará sus servicios en el domicilio del PATRÓN o en cualquier otro domicilio, centros de trabajo, clientes o instalaciones donde el Patrón tenga contratos de prestación de servicios de seguridad privada vigentes.

CUARTA. El Trabajador se obliga a cumplir estrictamente las consignas generales y particulares establecidas para cada servicio, procedimientos de acceso, control de visitantes, vigilancia perimetral y rondines. En caso de emergencias, su actuación se limitará estrictamente a activar los protocolos de seguridad pasiva y dar aviso inmediato a los cuerpos de auxilio públicos y central de operaciones.

QUINTA. El PATRÓN pagará al TRABAJADOR un salario ordinario de {datos['salario_semanal']} pesos semanales, cubriéndose los viernes de cada semana (incluye séptimos días y días festivos). Adicionalmente, el PATRÓN otorgará un Bono de Asistencia Semanal de $450.00 pesos y un Bono de Puntualidad Semanal de $450.00 pesos, condicionados al cumplimiento perfecto del 100% de asistencias y puntualidad validados mediante el sistema biométrico.

SEXTA. La duración máxima de la semana laboral será de 45 (cuarenta y cinco) horas en turnos rotativos (8x16, 12x12, 24x24 horas, etc.). Contará con 30 minutos intermedios para alimentos y reposo.

SEPTIMA. Los días de descanso semanal serán el domingo, sin perjuicio de que el PATRÓN modifique dichos días cuando las necesidades operativas del servicio lo requieran.

OCTAVA. Cuando el TRABAJADOR tenga más de un año de servicios, disfrutará de doce días de vacaciones anuales y una prima vacacional del 25%.

NOVENA. Serán días de descanso obligatorio los que señala el Artículo 74 de la Ley Federal del Trabajo, cubriéndose mediante roles operativos según los requerimientos de los clientes.

DECIMA. El PATRÓN pagará al TRABAJADOR un aguinaldo anual equivalente a 15 días de salario, a más tardar el 20 de diciembre de cada año.

DECIMA PRIMERA. El Patrón proporcionará los uniformes, gafetes, equipo de protección y herramientas necesarias, los cuales deberán utilizarse exclusivamente para fines laborales y devolverse al concluir la relación.

DECIMA SEGUNDA. El TRABAJADOR se obliga a cumplir estrictamente con el Reglamento Interior de Trabajo, mantener puntualidad, asistencia, permanencia en puesto hasta el relevo, estricta confidencialidad, portación correcta del uniforme y someterse a exámenes médicos y toxicológicos.

DECIMA TERCERA. CONFIDENCIALIDAD. El TRABAJADOR se obliga a guardar estricta confidencialidad sobre información, operaciones, clientes, estrategias y sistemas de seguridad de EL PATRÓN, tanto durante la vigencia del contrato como después de su terminación.

DECIMA CUARTA. DATOS PERSONALES. EL TRABAJADOR autoriza expresamente al PATRÓN para el tratamiento y resguardo de sus datos personales conforme a la Ley Federal de Protección de Datos Personales en Posesión de los Particulares.

DECIMA QUINTA. SEGURIDAD SOCIAL. EL PATRÓN se obliga a inscribir al TRABAJADOR ante el Instituto Mexicano del Seguro Social (IMSS) con el NSS {datos['nss']} y realizar las aportaciones correspondientes al INFONAVIT y SAR.

DECIMA SEXTA. CAPACITACIÓN. EL TRABAJADOR se obliga a participar en los cursos, talleres y programas de capacitación y adiestramiento organizados por EL PATRÓN conforme a los planes aprobados ante la STPS.

DECIMA SÉPTIMA. EXÁMENES MÉDICOS Y TOXICOLÓGICOS. EL TRABAJADOR acepta someterse a evaluaciones médicas, psicológicas y toxicológicas periódicas que EL PATRÓN solicite conforme a los requerimientos de control de confianza.

DECIMA OCTAVA. RESCISIÓN. Son causas de rescisión de la relación de trabajo, sin responsabilidad para EL PATRÓN, cualquiera de las señaladas en el Artículo 47 de la Ley Federal del Trabajo.

DECIMA NOVENA. FALTAS DE ASISTENCIA. Tres faltas de asistencia injustificadas en un periodo de treinta días constituirán causal de rescisión laboral sin responsabilidad patronal.

VIGÉSIMA. EQUIPO TÁCTICO Y ARMAS. En caso de portar equipos de comunicación o defensa autorizados, el TRABAJADOR es responsable de su cuidado y uso exclusivo conforme a los permisos vigentes de la empresa.

VIGÉSIMA PRIMERA. SUPERVISIÓN. El TRABAJADOR permitirá las supervisiones operativas sorpresa en su puesto asignado para verificar el correcto cumplimiento de sus funciones.

VIGÉSIMA SEGUNDA. MODIFICACIONES. Cualquier modificación a las condiciones generales de este contrato requerirá el consentimiento por escrito de ambas partes.

VIGÉSIMA TERCERA. REGLAMENTO INTERIOR. Las disposiciones del Reglamento Interior de Trabajo de la empresa forman parte integral de este contrato.

VIGÉSIMA CUARTA. DOMICILIOS. Las partes señalan como sus domicilios legales los indicados en el apartado de declaraciones de este instrumento.

VIGÉSIMA QUINTA. LEGISLACIÓN SUPLETORIA. Todo lo no previsto en el presente contrato se regirá por la Ley Federal del Trabajo y la Ley Federal de Seguridad Privada.

VIGÉSIMA SEXTA. JURISDICCIÓN. Para la interpretación y cumplimiento de este contrato, las partes se someten expresamente a la jurisdicción de los Tribunales Laborales competentes en el Estado de Nuevo León.

VIGÉSIMA SÉPTIMA. CONFORMIDAD. Enteradas las partes del contenido y alcance legal de cada una de las cláusulas del presente contrato, lo firman de su entera conformidad en Ciudad Benito Juárez, Nuevo León.

Se firma el presente contrato por duplicado en Ciudad Benito Juárez, Nuevo León.


EL PATRÓN
AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE S.A. DE C.V.
C. ABNER VELAZQUEZ MORALES



________________________________--



EL TRABAJADOR
{datos['nombre']}



________________________________--
"""

                for parrafo in texto_prueba.split("\n\n"):
                    if parrafo.strip():
                        doc.add_paragraph(parrafo.strip())

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                st.download_button(
                    label="📥 Clic aquí para descargar el Word listo",
                    data=buffer,
                    file_name=f"Contrato_Prueba_{datos['nombre'].replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

    elif menu_hr == "Generar Contrato Tiempo Indeterminado":
        st.header("📄 Generador de Contrato - Tiempo Indeterminado")

        if not st.session_state.empleados:
            st.warning(
                "⚠️ Primero registra un empleado en la sección 'Registro de Personal'."
            )
        else:
            nombres_empleados = [
                e["nombre"] for e in st.session_state.empleados
            ]
            emp_seleccionado = st.selectbox(
                "Selecciona al Trabajador", nombres_empleados
            )
            datos = next(
                e
                for e in st.session_state.empleados
                if e["nombre"] == emp_seleccionado
            )

            if st.button("📥 Descargar Contrato Indeterminado (Word)"):
                doc = Document()
                doc.add_heading(
                    "CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO INDETERMINADO",
                    level=1,
                )

                texto_indet = f"""CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO INDETERMINADO, QUE CELEBRAN, POR UNA PARTE, AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE, SOCIEDAD ANONIMA DE CAPITAL VARIABLE, REPRESENTADA EN ESTE ACTO POR EL C. ABNER VELAZQUEZ MORALES (EN LO SUCESIVO, EL "PATRÓN"), Y POR LA OTRA PARTE, POR SU PROPIO DERECHO, {datos['nombre']} (EN LO SUCESIVO, EL “TRABAJADOR”), DE CONFORMIDAD CON LOS ARTÍCULOS 20, 21, 24, 25, 35, 132, 134 Y DEMÁS RELATIVOS Y APLICABLES DE LA LEY FEDERAL DEL TRABAJO, AL TENOR DE LAS SIGUIENTES DECLARACIONES Y CLÁUSULAS:

D E C L A R A C I O N E S:

I. Declara el PATRÓN:
a) Ser una persona moral, debidamente constituida conforme a las leyes de la República Mexicana, según consta en la escritura pública número 6,948, pasada ante la fe del Notario Público número 127, con domicilio ubicado en Calle Santa Bárbara número 141, C. Asturias, Colonia Valle de Santa Isabel, C.P. 67256, Ciudad Benito Juárez, Nuevo León, Registro Federal de Contribuyentes AGI260413CK4 y tener como objeto social, entre otros, la prestación de servicios de seguridad privada, consistentes en la vigilancia, protección y resguardo de bienes muebles e inmuebles, así como de establecimientos comerciales, industriales, habitacionales y de servicios, mediante la utilización de recursos humanos debidamente capacitados.
b) Que, para dar cumplimiento a su objeto social, requiere de personal capacitado y con experiencia para ocupar el puesto de {datos['puesto']} para que realice las actividades consistentes, de manera enunciativa mas no limitativa, en: Vigilancia, protección y resguardo de bienes muebles e inmuebles, seguridad intramuros incluyendo el control de accesos y salidas, registro de personas y vehículos, realización de rondines y supervisión interna de instalaciones privadas.

II. Declara el TRABAJADOR:
a) Ser una persona física, de nacionalidad {datos['nacionalidad']}, de sexo {datos['sexo']}, con fecha de nacimiento el {datos['fecha_nacimiento']}, estado civil {datos['estado_civil']}, Clave Única de Registro de Población {datos['curp']}, Registro Federal de Contribuyentes {datos['rfc']} y Número de Seguridad Social (NSS) {datos['nss']}, con domicilio en {datos['domicilio']}.
b) Que cuenta con los conocimientos, habilidades y experiencia necesarios para prestar al PATRÓN los servicios del puesto encomendado.
c) Que está de acuerdo en prestar los servicios descritos en el presente contrato por tiempo indeterminado.

III. Declaran ambas partes:
a) Que cuentan con las facultades suficientes para la celebración del presente contrato y obligarse a los términos del mismo, reconociéndose mutuamente la personalidad con la que comparecen.

C L Á U S U L A S:

PRIMERA. El TRABAJADOR se obliga a prestar, bajo la dirección, dependencia y subordinación del PATRÓN, los servicios personales subordinados consistentes en las actividades del puesto de {datos['puesto']}. Las partes están de acuerdo en que los servicios mencionados se estipulan de manera enunciativa y no limitativa.

SEGUNDA. Se hace constar que el PATRÓN celebra el presente contrato fundado en las declaraciones del TRABAJADOR en el sentido de que cuenta con los requisitos y conocimientos necesarios para desempeñar adecuadamente las actividades inherentes al cargo para el que se le contrata.

TERCERA. El TRABAJADOR prestará sus servicios en el domicilio del PATRÓN o en cualquier otro domicilio en el que se ubiquen las oficinas, centros de trabajo, clientes o instalaciones donde el Patrón tenga contratos de prestación de servicios de seguridad privada vigentes.

CUARTA. El Trabajador se obliga a cumplir estrictamente las consignas generales y particulares establecidas para cada servicio, procedimientos de acceso, control de visitantes, vigilancia perimetral y rondines. En caso de emergencias, su actuación se limitará a activar los protocolos de seguridad pasiva y dar aviso a los cuerpos de auxilio y central de operaciones.

QUINTA. El PATRÓN pagará al TRABAJADOR un salario ordinario de {datos['salario_semanal']} pesos semanales, cubriéndose los viernes de cada semana. En este importe ya se incluye el pago correspondiente a los séptimos días y días festivos. Adicionalmente, el PATRÓN otorgará un Bono de Asistencia Semanal de $450.00 pesos y un Bono de Puntualidad Semanal de $450.00 pesos, condicionados al cumplimiento del 100% de asistencia y puntualidad validados biométricamente.

SEXTA. La duración máxima de la semana laboral será de 45 (cuarenta y cinco) horas en turnos rotativos (8x16, 12x12, 24x24 horas, etc.). Contará con 30 minutos intermedios para alimentos y reposo.

SEPTIMA. Los días de descanso semanal serán el domingo, sin perjuicio de las modificaciones por necesidades operativas del servicio.

OCTAVA. Cuando el TRABAJADOR tenga más de un año de servicios, disfrutará de doce días de vacaciones anuales y una prima vacacional del 25%.

NOVENA. Serán días de descanso obligatorio los señalados en el Artículo 74 de la Ley Federal del Trabajo.

DECIMA. El PATRÓN pagará al TRABAJADOR un aguinaldo anual equivalente a 15 días de salario, a más tardar el 20 de diciembre de cada año.

DECIMA PRIMERA. El Patrón proporcionará los uniformes, gafetes, equipo de protección y herramientas necesarias, los cuales deberán utilizarse exclusivamente para fines laborales y devolverse al concluir la relación.

DECIMA SEGUNDA. El TRABAJADOR se obliga a cumplir estrictamente con el Reglamento Interior de Trabajo, mantener puntualidad, asistencia, permanencia en puesto hasta el relevo, estricta confidencialidad, portación correcta del uniforme y someterse a exámenes médicos y toxicológicos.

DECIMA TERCERA. CONFIDENCIALIDAD. El TRABAJADOR se obliga a guardar estricta confidencialidad sobre información, operaciones, clientes, estrategias y sistemas de seguridad de EL PATRÓN, tanto durante la vigencia del contrato como después de su terminación.

DECIMA CUARTA. DATOS PERSONALES. EL TRABAJADOR autoriza expresamente al PATRÓN para el tratamiento y resguardo de sus datos personales conforme a la Ley Federal de Protección de Datos Personales en Posesión de los Particulares.

DECIMA QUINTA. SEGURIDAD SOCIAL. EL PATRÓN se obliga a inscribir al TRABAJADOR ante el Instituto Mexicano del Seguro Social (IMSS) con el NSS {datos['nss']} y realizar las aportaciones correspondientes al INFONAVIT y SAR.

DECIMA SEXTA. CAPACITACIÓN. El TRABAJADOR se obliga a participar en los cursos, talleres y programas de capacitación y adiestramiento organizados por EL PATRÓN conforme a los planes aprobados ante la STPS.

DECIMA SÉPTIMA. EXÁMENES MÉDICOS Y TOXICOLÓGICOS. EL TRABAJADOR acepta someterse a evaluaciones médicas, psicológicas y toxicológicas periódicas que EL PATRÓN solicite conforme a los requerimientos de control de confianza.

DECIMA OCTAVA. RESCISIÓN. Son causas de rescisión de la relación de trabajo, sin responsabilidad para EL PATRÓN, cualquiera de las señaladas en el Artículo 47 de la Ley Federal del Trabajo.

DECIMA NOVENA. FALTAS DE ASISTENCIA. Tres faltas de asistencia injustificadas en un periodo de treinta días constituirán causal de rescisión laboral sin responsabilidad patronal.

VIGÉSIMA. EQUIPO TÁCTICO Y ARMAS. En caso de portar equipos de comunicación o defensa autorizados, el TRABAJADOR es responsable de su cuidado y uso exclusivo conforme a los permisos vigentes de la empresa.

VIGÉSIMA PRIMERA. SUPERVISIÓN. El TRABAJADOR permitirá las supervisiones operativas sorpresa en su puesto asignado para verificar el correcto cumplimiento de sus funciones.

VIGÉSIMA SEGUNDA. MODIFICACIONES. Cualquier modificación a las condiciones generales de este contrato requerirá el consentimiento por escrito de ambas partes.

VIGÉSIMA TERCERA. REGLAMENTO INTERIOR. Las disposiciones del Reglamento Interior de Trabajo de la empresa forman parte integral de este contrato.

VIGÉSIMA CUARTA. DOMICILIOS. Las partes señalan como sus domicilios legales los indicados en el apartado de declaraciones de este instrumento.

VIGÉSIMA QUINTA. LEGISLACIÓN SUPLETORIA. Todo lo no previsto en el presente contrato se regirá por la Ley Federal del Trabajo y la Ley Federal de Seguridad Privada.

VIGÉSIMA SEXTA. JURISDICCIÓN. Para la interpretación y cumplimiento de este contrato, las partes se someten expresamente a la jurisdicción de los Tribunales Laborales competentes en el Estado de Nuevo León.

VIGÉSIMA SÉPTIMA. CONFORMIDAD. Enteradas las partes del contenido y alcance legal de cada una de las cláusulas del presente contrato, lo firman de su entera conformidad en Ciudad Benito Juárez, Nuevo León.

Se firma el presente contrato por duplicado en Ciudad Benito Juárez, Nuevo León.


EL PATRÓN
AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE S.A. DE C.V.
C. ABNER VELAZQUEZ MORALES



________________________________--



EL TRABAJADOR
{datos['nombre']}



________________________________--
"""

                for parrafo in texto_indet.split("\n\n"):
                    if parrafo.startX := parrafo.strip():
                        doc.add_paragraph(parrafo.strip())

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                st.download_button(
                    label="📥 Clic aquí para descargar el Word listo",
                    data=buffer,
                    file_name=f"Contrato_Indeterminado_{datos['nombre'].replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
