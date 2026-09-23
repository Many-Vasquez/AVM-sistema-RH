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
    page_title="AVM - Sistema de Seguridad Privada",
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

# --- BASES DE DATOS PERSISTENTES (CSV) ---
DB_FILE = "personal_avm.csv"
DB_ASISTENCIA = "asistencias_avm.csv"
DB_USUARIOS = "usuarios_avm.csv"
DB_AUDITORIA = "auditoria_avm.csv"


# Inicializar Administrador por defecto si no existe la BD de usuarios
def inicializar_usuarios():
    if not os.path.exists(DB_USUARIOS) or os.path.getsize(DB_USUARIOS) == 0:
        df_admin = pd.DataFrame(
            [
                {
                    "Usuario": "JOSE VASQUEZ",
                    "Password": "UCALLI123",
                    "Rol": "Administrador",
                }
            ]
        )
        df_admin.to_csv(DB_USUARIOS, index=False)


inicializar_usuarios()


def registrar_auditoria(usuario, accion, detalle):
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    nuevo_log = {
        "Fecha_Hora": ahora,
        "Usuario": usuario,
        "Acción": accion,
        "Detalle": detalle,
    }
    if os.path.exists(DB_AUDITORIA) and os.path.getsize(DB_AUDITORIA) > 0:
        df_log = pd.read_csv(DB_AUDITORIA)
        df_log = pd.concat([df_log, pd.DataFrame([nuevo_log])], ignore_index=True)
    else:
        df_log = pd.DataFrame([nuevo_log])
    df_log.to_csv(DB_AUDITORIA, index=False)


def cargar_usuarios():
    if os.path.exists(DB_USUARIOS) and os.path.getsize(DB_USUARIOS) > 0:
        try:
            return pd.read_csv(DB_USUARIOS).to_dict(orient="records")
        except Exception:
            return []
    return []


def guardar_usuarios(lista):
    pd.DataFrame(lista).to_csv(DB_USUARIOS, index=False)


def cargar_datos_empleados():
    if os.path.exists(DB_FILE) and os.path.getsize(DB_FILE) > 0:
        try:
            return pd.read_csv(DB_FILE).to_dict(orient="records")
        except Exception:
            return []
    return []


def guardar_datos_empleados(lista):
    pd.DataFrame(lista).to_csv(DB_FILE, index=False)


def cargar_datos_asistencias():
    if os.path.exists(DB_ASISTENCIA) and os.path.getsize(DB_ASISTENCIA) > 0:
        try:
            return pd.read_csv(DB_ASISTENCIA).to_dict(orient="records")
        except Exception:
            return []
    return []


def guardar_datos_asistencias(lista):
    pd.DataFrame(lista).to_csv(DB_ASISTENCIA, index=False)


# --- SISTEMA DE AUTENTICACIÓN (LOGIN) ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_actual = ""
    st.session_state.rol_actual = ""

logo_path = (
    "Imagen1 (1).png" if os.path.exists("Imagen1 (1).png") else "logo.png"
)

if not st.session_state.autenticado:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists(logo_path):
            st.image(logo_path, use_container_width=True)
        st.markdown(
            "<h1 style='text-align: center; color: #d4af37;'>AVM Grupo Integral de Seguridad Privada</h1>",
            unsafe_allow_html=True,
        )
        st.markdown(
            "<h3 style='text-align: center; color: #ffffff;'>Acceso al Sistema Operativo</h3>",
            unsafe_allow_html=True,
        )

        with st.form("form_login"):
            usuario_input = st.text_input("Usuario")
            password_input = st.text_input("Contraseña", type="password")
            btn_login = st.form_submit_button("🔑 Iniciar Sesión")

            if btn_login:
                usuarios_registrados = cargar_usuarios()
                user_match = next(
                    (
                        u
                        for u in usuarios_registrados
                        if u["Usuario"] == usuario_input
                        and str(u["Password"]) == password_input
                    ),
                    None,
                )

                if user_match:
                    st.session_state.autenticado = True
                    st.session_state.usuario_actual = user_match["Usuario"]
                    st.session_state.rol_actual = user_match["Rol"]
                    registrar_auditoria(
                        user_match["Usuario"],
                        "LOGIN",
                        "Inicio de sesión exitoso",
                    )
                    st.success("¡Acceso concedido! Cargando sistema...")
                    st.rerun()
                else:
                    st.error("Usuario o contraseña incorrectos.")
    st.stop()


# --- APLICACIÓN PRINCIPAL (UNA VEZ AUTENTICADO) ---
if os.path.exists(logo_path):
    st.sidebar.image(logo_path, width=150)

st.sidebar.markdown(
    f"👤 **Usuario:** {st.session_state.usuario_actual} (*{st.session_state.rol_actual}*)"
)
if st.sidebar.button("🚪 Cerrar Sesión"):
    registrar_auditoria(
        st.session_state.usuario_actual, "LOGOUT", "Cierre de sesión"
    )
    st.session_state.autenticado = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 🧭 Menú de Navegación", unsafe_allow_html=True)

opciones_menu = [
    "🏠 Inicio",
    "📊 Módulo Comercial (Cotizador)",
    "👥 Registro de Personal",
    "📥 Reporte de Personal (Excel)",
    "👆 Checador Biométrico de Huella",
    "📈 Reportes Métricos de Asistencia",
    "📄 Generación de Contratos",
]

# Si es Administrador, agregamos el panel de control maestro
if st.session_state.rol_actual == "Administrador":
    opciones_menu.append("🛡️ Panel de Administrador (Usuarios y Auditoría)")

menu = st.sidebar.radio("Seleccione el Módulo:", opciones_menu)


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


# --- 🏠 INICIO ---
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
            "<p style='text-align: center; color: #aaaaaa; font-size: 1.1rem;'>Sistema Operativo Centralizado</p>",
            unsafe_allow_html=True,
        )

# --- 📊 COTIZADOR ---
elif menu == "📊 Módulo Comercial (Cotizador)":
    st.header("📊 Módulo Comercial - Generador de Cotizaciones")
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
                registrar_auditoria(
                    st.session_state.usuario_actual,
                    "COTIZACION",
                    f"Generó cotización para {empresa_cliente}",
                )
                st.success("¡Datos de cotización cargados con éxito!")
            else:
                st.error(
                    "Por favor complete el nombre del cliente, cantidad y precio."
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
        run_dir = p_dir.add_run(
            "SANTA BÁRBARA NÚMERO 141, COLONIA VALLE DE SANTA ISABEL, C.P. 67256,\nCIUDAD BENITO JUÁREZ, NUEVO LEÓN"
        )
        run_dir.font.size = Pt(7.5)
        run_dir.font.color.rgb = COLOR_GRIS_TEXTO

        p_line = doc_cot.add_paragraph()
        r_line = p_line.add_run(
            "_________________________________________________________________________________"
        )
        r_line.font.size = Pt(8)
        r_line.font.color.rgb = COLOR_DORADO

        p_prop = doc_cot.add_paragraph()
        run_prop = p_prop.add_run("PROPUESTA ECONÓMICA DE SERVICIOS")
        run_prop.bold = True
        run_prop.font.size = Pt(11)
        run_prop.font.color.rgb = COLOR_DORADO

        p_datos = doc_cot.add_paragraph()
        if st.session_state.fecha_cot:
            p_datos.add_run(
                f"FECHA DE EMISIÓN:  {st.session_state.fecha_cot}\n"
            )
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
        r_h2_1 = h2_1.add_run("ANÁLISIS DE SITUACIÓN:")
        r_h2_1.font.size = Pt(10)
        r_h2_1.font.color.rgb = COLOR_DORADO
        p_analisis = doc_cot.add_paragraph(
            "Tras evaluar las necesidades de seguridad de su instalación, nuestra firma propone un esquema de Seguridad Proactiva. A diferencia de la vigilancia convencional, nuestro servicio se basa en la disuasión avanzada y la respuesta inmediata bajo los más altos estándares legales."
        )
        p_analisis.runs[0].font.size = Pt(9)
        p_analisis.runs[0].font.color.rgb = COLOR_NEGRO_SUAVE

        h2_2 = doc_cot.add_heading(level=2)
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
            "Control estricto de acceso peatonal y vehicular. Turno de 12 horas."
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
        r_h2_3 = h2_3.add_run("TÉRMINOS Y CONDICIONES COMERCIALES:")
        r_h2_3.font.size = Pt(10)
        r_h2_3.font.color.rgb = COLOR_DORADO
        terminos = [
            (
                "1. Responsabilidad Civil y Patronal:",
                "Obligaciones laborales y de seguridad social (IMSS, INFONAVIT) cubiertas.",
            ),
            (
                "2. Garantía de Continuidad:",
                "Cobertura al 100% con sustitución en menos de 90 minutos.",
            ),
            (
                "3. Confidencialidad Rigurosa:",
                "Contratos estrictos de secrecía para proteger al cliente.",
            ),
            (
                "4. Vigencia de la Propuesta:",
                "Validez de 15 días naturales a partir de su emisión.",
            ),
            (
                "5. Condiciones de Pago:",
                "Facturación mensual dentro de los primeros 5 días naturales.",
            ),
            (
                "6. Dias Festivos:",
                "Se cobran el doble del costo por dia.",
            ),
        ]
        for titulo, desc in terminos:
            p_term = doc_cot.add_paragraph()
            r_t = p_term.add_run(titulo + " ")
            r_t.bold = True
            r_t.font.size = Pt(8)
            r_t.font.color.rgb = COLOR_DORADO
            r_d = p_term.add_run(desc)
            r_d.font.size = Pt(8)
            r_d.font.color.rgb = COLOR_NEGRO_SUAVE

        p_pie = doc_cot.add_paragraph()
        p_pie.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_pie = p_pie.add_run(
            '"Nuestra estructura operativa garantiza la reducción del error humano mediante supervisión cruzada."'
        )
        r_pie.italic = True
        r_pie.font.size = Pt(8)
        r_pie.font.color.rgb = COLOR_DORADO

        buffer_cot = BytesIO()
        doc_cot.save(buffer_cot)
        buffer_cot.seek(0)
        st.success("¡Propuesta económica generada con éxito!")
        st.download_button(
            label="📥 Descargar Propuesta en Word",
            data=buffer_cot,
            file_name=f"Cotizacion_{st.session_state.empresa_cliente.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

# --- 👥 REGISTRO DE PERSONAL ---
elif menu == "👥 Registro de Personal":
    st.header("📝 Registro y Gestión de Personal / Guardias")
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

        submitted = st.form_submit_button("💾 Guardar Elemento")
        if submitted:
            if nombre and curp and nss:
                empleados = cargar_datos_empleados()
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
                empleados.append(nuevo_emp)
                guardar_datos_empleados(empleados)
                registrar_auditoria(
                    st.session_state.usuario_actual,
                    "ALTA PERSONAL",
                    f"Registró a {nombre} (NSS: {nss})",
                )
                st.success(f"¡Guardia {nombre} registrado correctamente!")
            else:
                st.error("Por favor complete Nombre, CURP y NSS.")

    st.markdown("---")
    st.subheader("⚙️ Modificar o Eliminar Personal")
    empleados = cargar_datos_empleados()
    if not empleados:
        st.info("No hay personal registrado.")
    else:
        nombres_registrados = [e["Nombre"] for e in empleados]
        emp_a_editar = st.selectbox("Seleccione Elemento", nombres_registrados)
        datos_actuales = next(e for e in empleados if e["Nombre"] == emp_a_editar)

        col_e1, col_e2 = st.columns(2)
        with col_e1:
            nuevo_nombre = st.text_input(
                "Modificar Nombre", value=datos_actuales["Nombre"]
            )
            nuevo_puesto = st.text_input(
                "Modificar Puesto", value=datos_actuales["Puesto"]
            )
            nuevo_nss = st.text_input(
                "Modificar NSS", value=datos_actuales["NSS"]
            )
        with col_e2:
            nuevo_salario = st.text_input(
                "Modificar Salario", value=datos_actuales["Salario Semanal"]
            )
            nuevo_domicilio = st.text_area(
                "Modificar Domicilio", value=datos_actuales["Domicilio"]
            )

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("🔄 Actualizar Datos"):
                for e in empleados:
                    if e["Nombre"] == emp_a_editar:
                        e["Nombre"] = nuevo_nombre
                        e["Puesto"] = nuevo_puesto
                        e["NSS"] = nuevo_nss
                        e["Salario Semanal"] = nuevo_salario
                        e["Domicilio"] = nuevo_domicilio
                guardar_datos_empleados(empleados)
                registrar_auditoria(
                    st.session_state.usuario_actual,
                    "MODIFICA PERSONAL",
                    f"Actualizó datos de {nuevo_nombre}",
                )
                st.success("¡Actualizado con éxito!")
                st.rerun()
        with col_btn2:
            if st.button("🗑️ Eliminar Elemento", type="primary"):
                empleados = [e for e in empleados if e["Nombre"] != emp_a_editar]
                guardar_datos_empleados(empleados)
                registrar_auditoria(
                    st.session_state.usuario_actual,
                    "ELIMINA PERSONAL",
                    f"Eliminó a {emp_a_editar}",
                )
                st.warning("Elemento eliminado.")
                st.rerun()

        st.subheader("📋 Plantilla Vigente")
        st.dataframe(pd.DataFrame(empleados), use_container_width=True)

# --- 📥 REPORTE DE PERSONAL (CSV/EXCEL) ---
elif menu == "📥 Reporte de Personal (Excel)":
    st.header("📥 Reporte General de Personal")
    empleados = cargar_datos_empleados()
    if not empleados:
        st.info("No hay registros.")
    else:
        df_excel = pd.DataFrame(empleados)
        st.dataframe(df_excel, use_container_width=True)
        st.download_button(
            "📊 Descargar Padrón (CSV)",
            df_excel.to_csv(index=False).encode("utf-8"),
            file_name="Personal_AVM.csv",
            mime="text/csv",
        )

# --- 👆 CHECADOR BIOMÉTRICO ---
elif menu == "👆 Checador Biométrico de Huella":
    st.header("👆 Terminal Biométrica de Asistencia")
    empleados = cargar_datos_empleados()
    if not empleados:
        st.warning("⚠️ No hay personal registrado.")
    else:
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            nombres = [e["Nombre"] for e in empleados]
            emp_sel = st.selectbox("Elemento", nombres)
            datos_emp = next(e for e in empleados if e["Nombre"] == emp_sel)
            tipo_mov = st.radio(
                "Movimiento", ["Entrada de Turno", "Salida de Turno"]
            )
        with col_b2:
            if st.button("🔴 ESCANEAR HUELLA DIGITAL", use_container_width=True):
                ahora = datetime.now()
                reg = {
                    "Fecha": ahora.strftime("%Y-%m-%d"),
                    "Nombre": datos_emp["Nombre"],
                    "NSS": datos_emp["NSS"],
                    "Movimiento": tipo_mov,
                    "Hora": ahora.strftime("%H:%M:%S"),
                    "Bono": "SÍ" if tipo_mov == "Salida de Turno" else "VALIDADO",
                }
                asistencias = cargar_datos_asistencias()
                asistencias.append(reg)
                guardar_datos_asistencias(asistencias)
                registrar_auditoria(
                    st.session_state.usuario_actual,
                    "ASISTENCIA",
                    f"Fichaje {tipo_mov} para {datos_emp['Nombre']}",
                )
                st.success(f"¡{tipo_mov} registrada!")

    asistencias = cargar_datos_asistencias()
    if asistencias:
        st.dataframe(pd.DataFrame(asistencias), use_container_width=True)

# --- 📈 REPORTES DE ASISTENCIA ---
elif menu == "📈 Reportes Métricos de Asistencia":
    st.header("📈 Auditoría y Reportes de Asistencia")
    asistencias = cargar_datos_asistencias()
    if not asistencias:
        st.info("Sin registros de asistencia.")
    else:
        df_asist = pd.DataFrame(asistencias)
        st.dataframe(df_asist, use_container_width=True)
        st.download_button(
            "📥 Exportar Asistencias",
            df_asist.to_csv(index=False).encode("utf-8"),
            file_name="Asistencias_AVM.csv",
            mime="text/csv",
        )

# --- 📄 CONTRATOS ---
elif menu == "📄 Generación de Contratos":
    st.header("📄 Generador de Contratos Laborales")
    empleados = cargar_datos_empleados()
    if not empleados:
        st.warning("⚠️ Sin personal registrado.")
    else:
        tipo_c = st.radio(
            "Tipo", ["Sujeto a Prueba (30 Días)", "Tiempo Indeterminado"]
        )
        emp_c = st.selectbox(
            "Trabajador", [e["Nombre"] for e in empleados]
        )
        datos = next(e for e in empleados if e["Nombre"] == emp_c)
        if st.button("📥 Generar Documento Word"):
            doc = Document()
            doc.add_heading(
                f"CONTRATO LABORAL - {datos['Nombre']}", level=1
            )
            doc.add_paragraph(
                f"Empresa: AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE S.A. DE C.V.\nPuesto: {datos['Puesto']}\nSalario: {datos['Salario Semanal']}"
            )
            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            registrar_auditoria(
                st.session_state.usuario_actual,
                "CONTRATO",
                f"Generó contrato para {datos['Nombre']}",
            )
            st.download_button(
                "📥 Descargar Word",
                buffer,
                file_name=f"Contrato_{datos['Nombre'].replace(' ', '_')}.docx",
            )

# --- 🛡️ PANEL DE ADMINISTRADOR MAESTRO ---
elif menu == "🛡️ Panel de Administrador (Usuarios y Auditoría)":
    if st.session_state.rol_actual != "Administrador":
        st.error(
            "⛔ Acceso restringido exclusivamente al Administrador Maestro."
        )
    else:
        st.header("🛡️ Panel de Control de Administrador")

        tab1, tab2 = st.tabs(
            ["👥 Gestión de Usuarios de la App", "📋 Bitácora de Auditoría Maestro"]
        )

        with tab1:
            st.subheader("Dar de Alta Nuevo Usuario para Codificar / Operar")
            with st.form("form_nuevo_usuario"):
                nuevo_user = st.text_input("Nombre de Usuario")
                nuevo_pass = st.text_input(
                    "Contraseña Temporal", type="password"
                )
                nuevo_rol = st.selectbox(
                    "Rol de Acceso", ["Operador", "Administrador"]
                )
                btn_crear = st.form_submit_button(
                    "➕ Registrar Nuevo Usuario"
                )

                if btn_crear:
                    if nuevo_user and nuevo_pass:
                        lista_u = cargar_usuarios()
                        if any(u["Usuario"] == nuevo_user for u in lista_u):
                            st.error("El usuario ya existe.")
                        else:
                            lista_u.append(
                                {
                                    "Usuario": nuevo_user,
                                    "Password": nuevo_pass,
                                    "Rol": nuevo_rol,
                                }
                            )
                            guardar_usuarios(lista_u)
                            registrar_auditoria(
                                st.session_state.usuario_actual,
                                "ALTA USUARIO",
                                f"Creó usuario '{nuevo_user}' con rol '{nuevo_rol}'",
                            )
                            st.success(
                                f"¡Usuario {nuevo_user} creado con éxito!"
                            )
                    else:
                        st.error("Complete todos los campos.")

            st.subheader("Usuarios con Acceso Autorizado")
            st.dataframe(
                pd.DataFrame(cargar_usuarios())[["Usuario", "Rol"]],
                use_container_width=True,
            )

        with tab2:
            st.subheader(
                "📋 Registro Central de Auditoría (Quién entró y qué editó)"
            )
            if os.path.exists(DB_AUDITORIA) and os.path.getsize(DB_AUDITORIA) > 0:
                df_auditoria = pd.read_csv(DB_AUDITORIA)
                st.dataframe(df_auditoria, use_container_width=True)

                st.download_button(
                    "📥 Descargar Respaldo Total de Auditoría (CSV)",
                    df_auditoria.to_csv(index=False).encode("utf-8"),
                    file_name=f"Auditoria_AVM_{datetime.now().strftime('%Y-%m-%d')}.csv",
                    mime="text/csv",
                )
            else:
                st.info("Aún no hay registros en la bitácora de auditoría.")
