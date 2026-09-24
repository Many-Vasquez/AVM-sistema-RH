from datetime import datetime, time
from io import BytesIO
import os
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
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

# --- CONSTANTES DE ARCHIVOS ---
ARCHIVO_PUNTOS = "puntos_trabajo.csv"

# --- FUNCIONES DE PERSISTENCIA Y CARGA DE DATOS ---
def cargar_datos(nombre_archivo, columnas=None):
    if os.path.exists(nombre_archivo):
        try:
            df = pd.read_csv(nombre_archivo)
            # Si se especifican columnas y el archivo está vacío, asegurarlas
            if columnas and df.empty:
                return pd.DataFrame(columns=columnas)
            return df
        except Exception:
            if columnas:
                return pd.DataFrame(columns=columnas)
            return pd.DataFrame()
    else:
        if columnas:
            return pd.DataFrame(columns=columnas)
        return pd.DataFrame()

def guardar_datos(df, nombre_archivo):
    df.to_csv(nombre_archivo, index=False)

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
DB_PUNTOS = "puntos_trabajo_avm.csv"


# Inicializar Administrador Maestro por defecto
def inicializar_usuarios():
    if not os.path.exists(DB_USUARIOS) or os.path.getsize(DB_USUARIOS) == 0:
        df_admin = pd.DataFrame(
            [
                {
                    "Usuario": "JOSE VASQUEZ",
                    "Password": "12345",
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

def cargar_datos_puntos(lista):
    if os.path.exists(DB_PUNTOS) and os.path.getsize(DB_PUNTOS) > 0:
        try:
            return pd.read_csv(DB_PUNTOS).to_dict("records")
        except Exception:
            return []
    return []

def guardar_datos_puntos(lista):
    pd.DataFrame(lista).to_csv(DB_PUNTOS, index=False)

puntos_act = cargar_datos_puntos([])

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
	"🏢 Catálogo Puntos de Trabajo",
	"📱 Terminal Móvil (Punto de Trabajo)",
    "👆 Checador Biométrico de Huella",
    "📈 Reportes Métricos de Asistencia",
    "📄 Generación de Contratos",
]

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

# ==========================================
# FUNCIÓN DE REGISTRO BIOMÉTRICO (WEBAUTHN)
# ==========================================
def registrar_huella_webauthn(empleado_nss):
    st.subheader("🔐 Registro de Credencial Biométrica (Celular Matriz)")
    st.markdown("Coloque el dedo en el sensor del teléfono para registrar la huella del elemento maestro.")

    biometric_html = f"""
    <div>
        <button id="bioBtn" style="background-color:#d4af37; color:black; padding:10px 20px; border:none; border-radius:5px; font-weight:bold; cursor:pointer;">
            Escanear Huella en Dispositivo
        </button>
        <p id="status" style="margin-top:10px; color:white;"></p>
    </div>
    
    <script>
    document.getElementById('bioBtn').onclick = async () => {{
        const statusEl = document.getElementById('status');
        try {{
            if (!window.PublicKeyCredential) {{
                statusEl.innerText = "Error: Este navegador o dispositivo no soporta biometría.";
                return;
            }}
            
            statusEl.innerText = "Escaneando huella... Por favor, use el sensor.";
            
            const publicKey = {{
                challenge: new Uint8Array([21, 31, 105, 43, 34, 45, 67, 89]),
                rp: {{ name: "AVM Grupo Integral de Seguridad Privada del Norte" }},
                user: {{
                    id: Uint8Array.from("{empleado_nss}", c => c.charCodeAt(0)),
                    name: "{empleado_nss}",
                    displayName: "Elemento AVM"
                }},
                pubKeyCredParams: [{{ alg: -7, type: "public-key" }}],
                timeout: 60000,
                attestation: "direct"
            }};

            const credential = await navigator.credentials.create({{ publicKey }});
            statusEl.innerText = "¡Huella registrada y validada con éxito!";
            console.log("Credencial creada:", credential.id);
            
        }} catch (error) {{
            statusEl.innerText = "Error o registro cancelado: " + error.message;
        }}
    }};
    </script>
    """
    components.html(biometric_html, height=150)

# ==========================================
# EJEMPLO DE INTEGRACIÓN EN EL MENÚ Y CONTROL DE ACCESO
# ==========================================
def gestionar_navegacion():
    # Verificamos si el usuario ha iniciado sesión y su rol es Administrador
    # (Asegúrate de adaptar 'st.session_state.get("role")' a la variable que uses para guardar el rol en tu login)
    rol_actual = st.session_state.get("role", "invitado")
    
    st.sidebar.title("Menú AVM Seguridad")
    
    # Opciones base para cualquier usuario autenticado
    opciones_menu = ["Asistencia", "Control Operativo"]
    
    # Si es Administrador, agregamos la opción exclusiva de Biometría
    if rol_actual == "Administrador":
        opciones_menu.append("Registro Biométrico Maestro")
        
    seleccion = st.sidebar.selectbox("Seleccione una opción", opciones_menu)
    
    # Lógica de las vistas
    if seleccion == "Asistencia":
        st.header("Control de Asistencia")
        st.write("Módulo de entradas y salidas de elementos.")
        
    elif seleccion == "Control Operativo":
        st.header("Panel Operativo")
        st.write("Gestión general de servicios y elementos.")
        
    elif seleccion == "Registro Biométrico Maestro" and rol_actual == "Administrador":
        st.header("Administración de Biometría Matriz")
        st.markdown("---")
        
        # Campo para ingresar el NSS del elemento al que se le registrará la huella
        nss_input = st.text_input("Ingrese el NSS (Número de Seguridad Social) del Elemento:")
        
        if nss_input:
            registrar_huella_webauthn(nss_input)
        else:
            st.info("Por favor, ingrese el NSS del elemento para habilitar el escaneo de huella.")

# Llamada principal a la navegación (colocar donde corresponda en tu app.py)
# gestionar_navegacion()

# --- 👥 REGISTRO DE PERSONAL ---
if menu == "Registro de Personal":
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

# --- 🏢 CATÁLOGO PUNTOS DE TRABAJO ---

elif menu == "🏢 Catálogo Puntos de Trabajo":
    st.title("🏢 Administración de Puntos de Trabajo (Clientes)")
    st.markdown("---")

    with st.form("form_punto"):
        nombre_punto = st.text_input("Nombre de la Instalación / Cliente")
        ubicacion = st.text_input("Dirección o Sector")
        btn_guardar_punto = st.form_submit_button("💾 Guardar Punto de Trabajo")

        if btn_guardar_punto:
            if nombre_punto:
                puntos = cargar_datos(
                    ARCHIVO_PUNTOS, ["NombrePunto", "Ubicacion", "FechaAlta"]
                )

                nueva_fila = pd.DataFrame([{
				"NombrePunto": nombre_punto, 
				"Ubicacion": ubicacion, 
				"FechaAlta": fecha_alta
				}])

                puntos = pd.concat([puntos, nueva_fila], ignore_index=True),
                guardar_datos(puntos, ARCHIVO_PUNTOS)
				st.success("¡Punto de trabajo guardado exitosamente!")
                }])
				
                guardar_datos(
                    ARCHIVO_PUNTOS,
                    puntos,
                    ["NombrePunto", "Ubicacion", "FechaAlta"],
                )
                registrar_auditoria(
                    st.session_state.usuario_actual,
                    "ALTA PUNTO",
                    f"Se creó el punto {nombre_punto}",
                )
                st.success(f"¡Punto de trabajo '{nombre_punto}' registrado con éxito!")
            else:
                st.warning("⚠️ Debes ingresar al menos el nombre del punto.")

    st.markdown("### 📋 Puntos de Trabajo Actuales")
    puntos_act = cargar_datos_puntos([])
    if puntos_act:
        st.dataframe(puntos_act, use_container_width=True)
    else:
        st.info("Aún no hay puntos de trabajo dados de alta.")
		
# --- 📱 TERMINAL MÓVIL (PUNTO DE TRABAJO) ---
elif menu == "📱 Terminal Móvil (Punto de Trabajo)":
    st.markdown(
        "<h2 style='text-align: center; color: #d4af37;'>Control Operativo en Campo</h2>",
        unsafe_allow_html=True,
    )

    empleados = cargar_datos_empleados()
    if not empleados:
        st.warning("⚠️ No hay personal registrado en el sistema central.")
    else:
        nombres = [e["Nombre"] for e in empleados]
        elemento_sel = st.selectbox("Seleccione su Nombre / Elemento", nombres)
        datos_elem = next(e for e in empleados if e["Nombre"] == elemento_sel)

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            btn_entrada = st.button(
                "🟢 REGISTRAR ENTRADA", use_container_width=True
            )
        with col_m2:
            btn_salida = st.button(
                "🔴 REGISTRAR SALIDA", use_container_width=True
            )

        if btn_entrada or btn_salida:
            ahora = datetime.now()
            tipo_mov = (
                "Entrada de Turno" if btn_entrada else "Salida de Turno"
            )
            reg = {
                "Fecha": ahora.strftime("%Y-%m-%d"),
                "Nombre": datos_elem["Nombre"],
                "NSS": datos_elem["NSS"],
                "Movimiento": tipo_mov,
                "Hora": ahora.strftime("%H:%M:%S"),
                "Punto de Trabajo": "Instalación Cliente",
            }
            asistencias = cargar_datos_asistencias()
            asistencias.append(reg)
            guardar_datos_asistencias(asistencias)
            registrar_auditoria(
                st.session_state.usuario_actual,
                "ASISTENCIA MÓVIL",
                f"Fichaje {tipo_mov} para {datos_elem['Nombre']}",
            )
            st.success(
                f"¡{tipo_mov} registrada correctamente a las {ahora.strftime('%H:%M:%S')}!"
            )

        st.markdown("---")
        st.markdown("### 🔄 Reportar Novedad o Cambio de Turno")
        with st.form("form_novedad_movil"):
            tipo_novedad = st.selectbox(
                "Tipo de Incidencia",
                [
                    "Cambio de Turno Solicitado",
                    "Doble Turno / Cobertura",
                    "Incidencia Operativa",
                ],
            )
            comentario = st.text_area("Detalle de la novedad")
            btn_enviar_nov = st.form_submit_button(
                "📤 Enviar a Central en Tiempo Real"
            )

            if btn_enviar_nov:
                registrar_auditoria(
                    datos_elem["Nombre"],
                    "NOVEDAD CAMPO",
                    f"[{tipo_novedad}] {comentario}",
                )
                st.success(
                    "¡Novedad enviada y registrada en la bitácora central!"
                )
				
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

            if "Prueba" in tipo_c:
                # Redacción exacta para Contrato Sujeto a Prueba (30 Días)
                texto_contrato = f"""CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO INDETERMINADO, SUJETO A UN PERIODO DE PRUEBA, QUE CELEBRAN, POR UNA PARTE, AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE, SOCIEDAD ANONIMA DE CAPITAL VARIABLE, REPRESENTADA EN ESTE ACTO POR EL C. ABNER VELAZQUEZ MORALES (EN LO SUCESIVO, EL "PATRÓN"), Y POR LA OTRA PARTE, POR SU PROPIO DERECHO, {datos['Nombre']} (EN LO SUCESIVO, EL “TRABAJADOR”), DE CONFORMIDAD CON LOS ARTÍCULOS 20, 21, 24, 25, 35, 39-A, 39-B, 132, 134 Y DEMÁS RELATIVOS Y APLICABLES DE LA LEY FEDERAL DEL TRABAJO, AL TENOR DE LAS SIGUIENTES DECLARACIONES Y CLÁUSULAS:

D E C L A R A C I O N E S:

I. Declara el PATRÓN:
	a) Ser una persona moral, debidamente constituida conforme a las leyes de la República Mexicana, según consta en la escritura pública número 6,948, pasada ante la fe del Notario Público número 127, con domicilio ubicado en Calle Santa Bárbara número 141, C. Asturias, Colonia Valle de Santa Isabel, C.P. 67256, Ciudad Benito Juárez, Nuevo León, Registro Federal de Contribuyentes AGI260413CK4 y tener como objeto social, entre otros, la prestación de servicios de seguridad privada, consistentes en la vigilancia, protección y resguardo de bienes muebles e inmuebles, así como de establecimientos comerciales, industriales, habitacionales y de servicios, mediante la utilización de recursos humanos debidamente capacitados. La prestación de servicios de seguridad intramuros, incluyendo el control de accesos y salidas, registro de personas y vehículos, realización de rondines y supervisión interna de instalaciones privadas, con la finalidad de prevenir actos delictivos, riesgos o siniestros. La administración y control de acceso por medio de personal, así como la implementación de protocolos de seguridad, supervisión operativa y control de ingreso y egreso en todo tipo de instalaciones.
	b) Que, para dar cumplimiento al objeto social o profesión consignado en la declaración anterior, requiere de personal capacitado y con experiencia para ocupar el puesto de {datos['Puesto']} para que realice las actividades consistentes, de manera enunciativa mas no limitativa, en:
	Vigilancia, protección y resguardo de bienes muebles e inmuebles, así como en establecimientos comerciales, industriales, habitaciones. Seguridad intramuros incluyendo el control de accesos y salidas, registro de personas y vehículos, realización de rondines y supervisión interna de instalaciones privadas, operación y monitoreo de sistemas de seguridad electrónica, sistemas de alarma, circuito cerrado de televisión, sistemas de control de accesos, sistemas de rastreo satelital y demás tecnologías relacionadas con la seguridad.

II. Declara el TRABAJADOR:
	a) Ser una persona física, de nacionalidad {datos['Nacionalidad']}, de sexo {datos['Sexo']}, con fecha de nacimiento el {datos.get('Fecha de Nacimiento', '[No especificada]')}, estado civil {datos.get('Estado Civil', '[No especificado]')}, Clave Única de Registro de Población {datos['CURP']} y Registro Federal de Contribuyentes {datos['RFC']}, con domicilio en {datos['Domicilio']}.
	b) Que cuenta con los conocimientos, habilidades y experiencia necesarios para prestar al PATRÓN los servicios mencionados en el inciso b) de la Declaración I.
	c) Que está de acuerdo en prestar los servicios descritos en el presente contrato, sujeto a un periodo de prueba de 30 (treinta) días, según se estipula más adelante.

III. Declaran ambas partes:
	a) Que cuentan con las facultades suficientes para la celebración del presente contrato y obligarse a los términos de este, reconociéndose mutuamente la personalidad con la que comparecen.

En consideración a las Declaraciones que anteceden, las partes otorgan las siguientes:

C L Á U S U L A S:

	PRIMERA. El presente contrato se celebra por TIEMPO INDETERMINADO, quedando sujeto "EL TRABAJADOR" a un PERIODO DE PRUEBA DE 1 MES (30 DÍAS) contados a partir de la fecha de firma del presente contrato, con fundamento en el párrafo segundo del Artículo 39-A de la Ley Federal del Trabajo, toda vez que el puesto a desempeñar requiere de labores técnicas, operativas y/o conocimientos especializados en materia de seguridad privada, prevención de riesgos y manejo de equipos tácticos/tecnológicos.
Durante dicho periodo de prueba, "EL PATRÓN" evaluará si "EL TRABAJADOR" cumple con los requisitos, conocimientos y aptitudes necesarios para el puesto. De no acreditarlos a satisfacción de "EL PATRÓN" mediante la evaluación del Comité Mixto de Productividad, Capacitación y Adiestramiento, podrá dar por terminada la relación de trabajo en cualquier momento, sin responsabilidad alguna para la Empresa y sin obligación de pagar indemnización constitucional alguna, procediendo únicamente al pago del finiquito proporcional de las prestaciones devengadas.

	SEGUNDA. Se hace constar que el PATRÓN celebra el presente contrato fundado en las declaraciones del TRABAJADOR en el sentido de que cuenta con los requisitos y conocimientos necesarios para desempeñar adecuadamente las actividades inherentes al cargo para el que se le contrata. Al término del periodo de prueba, de no acreditar el TRABAJADOR que satisface los requisitos y conocimientos necesarios para desarrollar las labores, a juicio del patrón, tomando en cuenta la opinión de la Comisión Mixta de Productividad, Capacitación y Adiestramiento en los términos de la Ley Federal del Trabajo, así́ como la naturaleza de la categoría o puesto, se dará́ por terminada la relación de trabajo, sin responsabilidad para el PATRÓN.

	TERCERA. El TRABAJADOR prestará sus servicios en el domicilio del PATRÓN o en cualquier otro domicilio en el que se ubiquen las oficinas o locales del mismo. El TRABAJADOR, manifiesta desde este momento su conformidad con cualquier eventual cambio en el lugar de la prestación de sus servicios. El Trabajador acepta que, debido a la naturaleza de los servicios de seguridad privada, podrá ser asignado temporal o permanentemente a distintos centros de trabajo, clientes, instalaciones o ubicaciones donde el Patrón tenga contratos de prestación de servicios. Dichos cambios no constituirán modificación unilateral de las condiciones de trabajo siempre que se respeten los derechos laborales del trabajador.

	CUARTA. El Trabajador se obliga a cumplir estrictamente las consignas generales y particulares establecidas para cada servicio, incluyendo procedimientos de acceso, control de visitantes, vigilancia perimetral, rondines y reportes.  De conformidad con el Reglamento Interior de Trabajo de la Empresa, las funciones del TRABAJADOR están limitadas estrictamente a la prevención, vigilancia y control de accesos. En caso de emergencias (tales como siniestros, robos en proceso o accidentes), la intervención del TRABAJADOR se limitará de manera enunciativa más no limitativa a: activar los protocolos de seguridad pasiva, dar aviso inmediato a los cuerpos de auxilio públicos (policía, bomberos, ambulancias), reportar a la central de operaciones de El PATRÓN y auxiliar en la evacuación segura del personal. Queda estrictamente prohibido realizar acciones de confrontación o tácticas que pongan en riesgo su integridad física o la de terceros. La omisión injustificada en el cumplimiento de estas consignas será sancionada conforme a la Ley Federal del Trabajo y al Reglamento Interior de Trabajo.

	QUINTA. El PATRÓN pagará al TRABAJADOR, por los servicios prestados de conformidad con este contrato, un salario ordinario de {datos['Salario Semanal']} pesos semanales, el cual se cubrirá los viernes de cada semana. En este importe ya se encuentra incluido el pago correspondiente a los séptimos días (días de descanso semanal) y los días festivos de descanso obligatorio en términos de los artículos 69 y 74 de la Ley Federal del Trabajo. 
De los Bonos de Asistencia y Puntualidad: Adicionalmente al salario ordinario, el PATRÓN otorgará al TRABAJADOR un Bono de Asistencia Semanal por la cantidad de $450.00 pesos (cuatrocientos cincuenta pesos 00/100 M.N.) y un Bono de Puntualidad Semanal por la cantidad de $450.00 pesos (cuatrocientos cincuenta pesos 00/100 M.N.).
El TRABAJADOR queda estrictamente obligado al cumplimiento del 100% de sus asistencias y de sus horarios de entrada durante la semana correspondiente para devengar dichos conceptos. Las partes acuerdan que el nacimiento del derecho a recibir estos bonos está condicionado estrictamente al cumplimiento perfecto de la asistencia y puntualidad; por lo tanto, en caso de que el TRABAJADOR incurra en una sola falta de asistencia (justificada o injustificada) o en un solo retardo durante el periodo semanal, no se generará ni se pagará el bono correspondiente a la falta incurrida (ya sea de asistencia, de puntualidad, o ambos), sin responsabilidad alguna para el PATRÓN.
En caso de que el día de pago sea de descanso obligatorio o festivo bancario, el depósito se efectuará el día hábil inmediato anterior. El pago del salario y de las prestaciones que correspondan se realizará exclusivamente mediante transferencia electrónica de fondos a la cuenta bancaria institucional que EL PATRÓN apertura a nombre del TRABAJADOR, o a la cuenta que este designe por escrito. Los costos de apertura y manejo de cuenta correrán por cuenta de la Empresa. Al importe del salario y bonos se le realizarán las deducciones legales de impuestos y Seguridad Social correspondientes. 
	Por su parte, el TRABAJADOR se obliga, en cualquier caso, a firmar el recibo correspondiente por los pagos efectuados. Si, por alguna razón, éste no firmara el recibo, las partes aceptan que el simple depósito bancario produce efecto liberatorio de pago para el PATRÓN.  

	SEXTA. La duración máxima de la semana laboral será de 45 (cuarenta y cinco) horas, distribuidas de lunes a sábado de cada semana, de conformidad con lo dispuesto por el segundo párrafo del artículo 59 de la Ley Federal del Trabajo. 
Debido a la naturaleza especializada de las actividades de seguridad y vigilancia que presta la Empresa, y para garantizar la continuidad y la cobertura ininterrumpida de los servicios contratados por nuestros clientes, las partes acuerdan expresamente que las jornadas y horarios de trabajo no serán fijos ni permanentes. El TRABAJADOR prestará sus servicios bajo esquemas de turnos rotativos (tales como 8x16 horas, 12x12 horas, 24x24 horas, o los esquemas que operativamente se requieran) , según el rol que le sea comunicado oportunamente por su superior inmediato o supervisor de zona. 
Dentro de la jornada continua, el TRABAJADOR dispondrá de un lapso de 30 (treinta) minutos intermedios para tomar alimentos y reposar, el cual será considerado como tiempo efectivo de trabajo y se adaptará de forma flexible a las necesidades de cada servicio ; durante este periodo, el trabajador podrá utilizar las sillas o asientos destinados para tal efecto de acuerdo con el artículo 132 fracción V de la Ley Federal del Trabajo. El PATRÓN podrá modificar en cualquier tiempo el horario, rol y la rotación de turnos conforme a las necesidades operativas de la Empresa. 
El TRABAJADOR no laborará tiempo extra en su jornada normal, ni durante días de descanso, salvo previa orden expresa y por escrito emitida por el representante del PATRÓN ; orden sin la cual no se reconocerá ni se pagará tiempo extraordinario alguno. 

	SEPTIMA. Las partes convienen en que los días de descanso semanal serán el domingo, sin perjuicio de que el PATRÓN modifique dichos días de descanso semanal cuando las necesidades del servicio así lo requieran. 

	OCTAVA. Cuando el TRABAJADOR tenga más de un año de servicios, disfrutará de doce días de vacaciones anuales en los términos y condiciones que establece la Ley Federal del Trabajo. Adicionalmente, el PATRÓN pagará al TRABAJADOR una prima vacacional del 25% sobre el salario que le corresponda por sus días de vacaciones, en términos de lo dispuesto por el Artículo 80 de la Ley Federal del Trabajo.

	NOVENA. Serán días de descanso obligatorio los que señala el Artículo 74 de la Ley Federal del Trabajo. Atendiendo a la naturaleza especializada de los servicios de seguridad y vigilancia de la Empresa, la cobertura de estos días estará sujeta a los roles operativos cambiantes y requerimientos de los clientes asignados por el supervisor de zona, obligándose el TRABAJADOR a prestar sus servicios si la operación lo requiere, previo pago de las compensaciones legales aplicables.

	DECIMA. El PATRÓN pagará al TRABAJADOR un aguinaldo anual, equivalente a 15 días de salario, en los términos que establece el Artículo 87 de la Ley Federal del Trabajo, mismo que deberá cubrirse a más tardar el 20 de diciembre de cada año.

	DECIMA PRIMERA. EQUIPO Y UNIFORMES. El Patrón proporcionará los uniformes, gafetes, equipo de protección y herramientas necesarias para el desempeño de sus funciones. Los radios, teléfonos, cámaras corporales, dispositivos electrónicos y demás equipos proporcionados por la empresa deberán utilizarse exclusivamente para fines laborales.
Queda prohibido alterar, modificar, cambiar, dañar, prestar o utilizar dichos equipos para fines personales. El Trabajador se obliga a utilizarlos adecuadamente y devolverlos al concluir la relación laboral.

	DECIMA SEGUNDA. OBLIGACIONES DEL TRABAJADOR. Además de las obligaciones previstas en el artículo 134 de la Ley Federal del Trabajo, el TRABAJADOR se obliga estrictamente a cumplir con lo establecido en el Reglamento Interior de Trabajo de la Empresa, comprometiéndose de manera enunciativa más no limitativa a: 
Cumplimiento de Instrucciones y Consignas: Acatar con eficacia, cuidado y esmero las instrucciones de trabajo, órdenes patronales y las consignas específicas o particulares establecidas para el puesto o servicio asignado. 
Puntualidad y Asistencia: Presentarse puntualmente a sus labores respetando los horarios y roles de turnos asignados, registrando personalmente su entrada y salida en los controles (tarjeta, bitácora o lector) que determine el PATRÓN. 
Permanencia y Protocolo de Relevo: Permanecer de forma estrictamente personal en su puesto de vigilancia y control de accesos hasta que se presente físicamente su relevo y se realice la entrega formal de la bitácora y equipo; reconociendo que el abandono del puesto sin autorización expresa del supervisor constituirá una falta grave. 
Estricta Confidencialidad y Reserva: Guardar absoluta reserva y discreción sobre los asuntos de la Empresa y sus clientes. Queda estrictamente prohibido fotografiar, videograbar, reproducir, extraer o difundir por cualquier medio (incluyendo redes sociales o WhatsApp) el contenido de bitácoras, controles de acceso, sistemas de monitoreo o pantallas de CCTV. 
Uso y Portación del Uniforme: Mantener una imagen personal aseada y profesional, portando correctamente el uniforme completo, limpio, fajado y con el gafete de identificación visible durante toda su jornada laboral, como medida de confianza hacia los clientes. 
Inspección y Conservación de Equipo: Revisar, inspeccionar y conservar en buen estado las herramientas, útiles, uniformes y el equipo táctico o de comunicación (radios, fornituras, linternas) proporcionados por el PATRÓN; reportando inmediatamente cualquier desperfecto y absteniéndose de utilizarlos para fines personales. 
Reporte de Incidencias: Informar de manera inmediata a su supervisor de zona o a la central de operaciones sobre cualquier incidente de seguridad, anomalía, siniestro o situación de riesgo detectada en las instalaciones. 
Conducta Profesional: Mantener en todo momento un trato respetuoso, digno y profesional con los clientes, visitantes, proveedores y compañeros de trabajo. 
Exámenes de Control y Confianza: Someterse a los exámenes médicos, psicométricos, de alcoholemia y toxicológicos aleatorios o periódicos que determine la Empresa o las autoridades competentes para prevenir riesgos de trabajo; aceptando que la negativa a realizárselos será causa de rescisión inmediata de la relación de trabajo. 
Medidas Preventivas de Seguridad e Higiene: Observar y acatar rigurosamente todas las medidas preventivas, higiénicas y de seguridad que acuerden las autoridades y las que indique el PATRÓN para salvaguardar su integridad física, la de sus compañeros y la de los bienes resguardados. 

	DECIMA TERCERA. El trabajador deberá elaborar y entregar oportunamente los reportes, bitácoras, formatos de novedades y demás documentos operativos requeridos por la empresa. La falsificación de información o la omisión deliberada de hechos relevantes constituirá falta grave.

	DECIMA CUARTA. El TRABAJADOR se obliga a obedecer estrictamente las normas de trabajo fijadas por el PATRÓN y a respetar la organización jerárquica que la misma tiene establecidas o en el futuro establezca, así como a cumplir con todas las obligaciones que naturalmente deriven de este contrato y de los servicios que debe prestar. 

	DECIMA QUINTA. El TRABAJADOR reconoce que son propiedad exclusiva del PATRÓN y/o sus clientes todos los documentos e información que se le proporcionen con motivo de la relación de trabajo, así como los que el propio PATRÓN prepare o formule en relación o conexión con sus servicios, por lo que se obliga a conservarlos en buen estado y a entregarlos al PATRÓN en el momento en que éste lo requiera o bien al terminar el presente contrato, por el motivo que sea.

	DECIMA SEXTA. El TRABAJADOR se obliga a devolver, a satisfacción del PATRÓN, los instrumentos, equipos y materiales que le fueren proporcionados para el desempeño de sus funciones en el momento en que éste lo requiera o al término del presente acuerdo de voluntades por el motivo que fuere.
Si el EMPLEADO dejare de cumplir con lo establecido en la presente Clausula, quedará sujeto a la responsabilidad civil por los daños o perjuicios que causare al PATRÓN como dueño y propietario de estas herramientas de trabajo, así como las sanciones de carácter penal a que por ello se hiciere acreedor.     

	DECIMA SEPTIMA. CONFIDENCIALIDAD. 
El TRABAJADOR, en cumplimiento a la fracción XIII del artículo 134 de la Ley Federal del Trabajo, se obliga a no divulgar ninguno de los aspectos de los negocios del PATRÓN, información de clientes, vulnerabilidades de clientes, manuales y datos personales a los que tenga acceso con motivo de su trabajo.ni datos personales a terceras personas, verbalmente o por escrito, directa o indirectamente, información alguna sobre los sistemas o actividades de cualquier clase que observe el PATRÓN. La revelación de esta información a terceros será causa de rescisión de la relación laboral.
Si el TRABAJADOR dejare de cumplir con las disposiciones de esta Cláusula, quedará sujeto a la responsabilidad civil por daños y perjuicios que cause al PATRÓN y a las sanciones penales que marca la ley a que se haga acreedor.

	DECIMA OCTAVA. El TRABAJADOR se obliga a cumplir con las disposiciones legales en materia de protección de datos personales y a resguardar adecuadamente cualquier información a la que tenga acceso durante la prestación de sus servicios.

	DECIMA NOVENA. Conforme a lo dispuesto por la Fracción X del Artículo 134 de la Ley Federal del Trabajo, el TRABAJADOR se someterá a los exámenes médicos que ordene el PATRÓN, en la inteligencia de que el facultativo que los practique será designado y retribuido por éste mismo. 

	VIGESIMA. Para todo lo relacionado con riesgos de trabajo y enfermedades o accidentes no profesionales, se estará a lo dispuesto por la Ley del Seguro Social y sus Reglamentos, para lo cual el PATRÓN inscribirá oportunamente al TRABAJADOR ante el Instituto Mexicano del Seguro Social, cubriéndose las cuotas por ambas partes, en los términos que consigna la citada Ley.
En tal virtud, el único documento válido para justificar faltas de asistencia derivadas de incapacidad por enfermedad general o profesional será el certificado y/o incapacidad que expida el Instituto Mexicano del Seguro Social.

	VIGESIMA PRIMERA. El PATRÓN proporcionará capacitación y adiestramiento al TRABAJADOR, conforme a los planes y programas establecidos, o que se establezcan de acuerdo con las disposiciones de la Ley Federal del Trabajo, comprometiéndose el TRABAJADOR a dedicar el tiempo y esfuerzo necesarios para lograr la mejor eficiencia en dicha capacitación.

	VIGESIMA SEGUNDA. Las partes convienen que, en lo no previsto por el presente contrato, se sujetarán a las disposiciones de la Ley Federal del Trabajo y de la Ley del Seguro Social y sus Reglamentos

	VIGESIMA TERCERA. CAUSAS DE RESCISIÓN:
Serán causas de rescisión sin responsabilidad para el Patrón las previstas en el artículo 47 de la Ley Federal del Trabajo y demás disposiciones aplicables.

	VIGESIMA CUARTA. Salvo autorización expresa del supervisor o cliente, queda restringido el uso de teléfonos celulares, audífonos, tabletas u otros dispositivos electrónicos durante la prestación del servicio cuando ello afecte la vigilancia o seguridad del puesto asignado.

	VIGESIMA QUINTA. REGLAMENTO INTERIOR DE TRABAJO.
El Trabajador manifiesta conocer y aceptar el Reglamento Interior de Trabajo de la empresa, obligándose a cumplirlo en todos sus términos y demás disposiciones internas de la empresa.

	VIGESIMA SEXTA. LEGISLACIÓN APLICABLE.
Para todo lo no previsto en este contrato serán aplicables la Ley Federal del Trabajo, la legislación laboral vigente, la normativa de seguridad social y las disposiciones aplicables a la prestación de servicios de seguridad privada en el Estado de Nuevo León.

	VIGÉSIMA SEPTIMA. Cualquier modificación que se haga al presente contrato deberá constar por escrito y deberá ser firmada de conformidad por ambas partes.

Leído que lo fue íntegramente el presente contrato y enteradas las partes de su contenido y alcancel legal, el TRABAJADOR y el PATRÓN lo ratificaron y firmaron de conformidad, por duplicado, ante dos testigos, en la calle Santa Barbara número 141, Colonia Valle de Santa Isabel, C.P. 67256, Ciudad Benito Juarez, Nuevo León, el _____ de _________________ de 202__, quedando un original en poder del PATRÓN y otro en poder del TRABAJADOR.

EL PATRON
AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE S.A. DE C.V.
C. ABNER VELAZQUEZ MORALES
Representante legal
 | EL TRABAJADOR
{datos['Nombre']}
Por sus propios derechos

TESTIGO
(Nombre)
(Dirección)
 | TESTIGO
(Nombre)
(Dirección)
"""
            else:
                # Redacción exacta para Contrato por Tiempo Indeterminado (Sin periodo de prueba)
                texto_contrato = f"""CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO INDETERMINADO, QUE CELEBRAN, POR UNA PARTE, AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE, SOCIEDAD ANONIMA DE CAPITAL VARIABLE, REPRESENTADA EN ESTE ACTO POR EL C. ABNER VELAZQUEZ MORALES (EN LO SUCESIVO, EL "PATRÓN"), Y POR LA OTRA PARTE, POR SU PROPIO DERECHO, {datos['Nombre']} (EN LO SUCESIVO, EL “TRABAJADOR”), DE CONFORMIDAD CON LOS ARTÍCULOS 20, 21, 24, 25, 35, 39-A, 39-B, 132, 134 Y DEMÁS RELATIVOS Y APLICABLES DE LA LEY FEDERAL DEL TRABAJO, AL TENOR DE LAS SIGUIENTES DECLARACIONES Y CLÁUSULAS:

D E C L A R A C I O N E S:

I. Declara el PATRÓN:
	a) Ser una persona moral, debidamente constituida conforme a las leyes de la República Mexicana, según consta en la escritura pública número 6,948, pasada ante la fe del Notario Público número 127, con domicilio ubicado en Calle Santa Bárbara número 141, C. Asturias, Colonia Valle de Santa Isabel, C.P. 67256, Ciudad Benito Juárez, Nuevo León, Registro Federal de Contribuyentes AGI260413CK4 y tener como objeto social, entre otros, la prestación de servicios de seguridad privada, consistentes en la vigilancia, protección y resguardo de bienes muebles e inmuebles, así como de establecimientos comerciales, industriales, habitacionales y de servicios, mediante la utilización de recursos humanos debidamente capacitados. La prestación de servicios de seguridad intramuros, incluyendo el control de accesos y salidas, registro de personas y vehículos, realización de rondines y supervisión interna de instalaciones privadas, con la finalidad de prevenir actos delictivos, riesgos o siniestros. La administración y control de acceso por medio de personal, así como la implementación de protocolos de seguridad, supervisión operativa y control de ingreso y egreso en todo tipo de instalaciones.
	b) Que, para dar cumplimiento al objeto social o profesión consignado en la declaración anterior, requiere de personal capacitado y con experiencia para ocupar el puesto de {datos['Puesto']} para que realice las actividades consistentes, de manera enunciativa mas no limitativa, en:
	Vigilancia, protección y resguardo de bienes muebles e inmuebles, así como en establecimientos comerciales, industriales, habitaciones. Seguridad intramuros incluyendo el control de accesos y salidas, registro de personas y vehículos, realización de rondines y supervisión interna de instalaciones privadas, operación y monitoreo de sistemas de seguridad electrónica, sistemas de alarma, circuito cerrado de televisión, sistemas de control de accesos, sistemas de rastreo satelital y demás tecnologías relacionadas con la seguridad.

II. Declara el TRABAJADOR:
	a) Ser una persona física, de nacionalidad {datos['Nacionalidad']}, de sexo {datos['Sexo']}, con fecha de nacimiento el {datos.get('Fecha de Nacimiento', '[No especificada]')}, estado civil {datos.get('Estado Civil', '[No especificado]')}, Clave Única de Registro de Población {datos['CURP']} y Registro Federal de Contribuyentes {datos['RFC']}, con domicilio en {datos['Domicilio']}.
	b) Que cuenta con los conocimientos, habilidades y experiencia necesarios para prestar al PATRÓN los servicios mencionados en el inciso b) de la Declaración I.
	c) Que está de acuerdo en prestar los servicios descritos en el presente contrato.

III. Declaran ambas partes:
	a) Que cuentan con las facultades suficientes para la celebración del presente contrato y obligarse a los términos del mismo, reconociéndose mutuamente la personalidad con la que comparecen.

En consideración a las Declaraciones que anteceden, las partes otorgan las siguientes:

C L Á U S U L A S:

	PRIMERA. El TRABAJADOR se obliga a prestar, bajo la dirección, dependencia y subordinación del PATRÓN, los servicios personales descritos en el inciso b) de la Declaración I, con el puesto de {datos['Puesto']}. Las partes están de acuerdo en que los servicios mencionados anteriormente se estipulan de manera enunciativa y no limitativa, por lo que el TRABAJADOR se obliga a desempeñar todas las labores anexas y conexas que le ordene el PATRÓN.

	SEGUNDA. Se hace constar que el PATRÓN celebra el presente contrato fundado en las declaraciones del TRABAJADOR en el sentido de que cuenta con los requisitos y conocimientos necesarios para desempeñar adecuadamente las actividades inherentes al cargo para el que se le contrata. Al término del periodo de prueba, de no acreditar el TRABAJADOR que satisface los requisitos y conocimientos necesarios para desarrollar las labores, a juicio del patrón, tomando en cuenta la opinión de la Comisión Mixta de Productividad, Capacitación y Adiestramiento en los términos de la Ley Federal del Trabajo, así́ como la naturaleza de la categoría o puesto, se dará́ por terminada la relación de trabajo, sin responsabilidad para el PATRÓN.

	TERCERA. El TRABAJADOR prestará sus servicios en el domicilio del PATRÓN o en cualquier otro domicilio en el que se ubiquen las oficinas o locales del mismo. El TRABAJADOR, manifiesta desde este momento su conformidad con cualquier eventual cambio en el lugar de la prestación de sus servicios. El Trabajador acepta que, debido a la naturaleza de los servicios de seguridad privada, podrá ser asignado temporal o permanentemente a distintos centros de trabajo, clientes, instalaciones o ubicaciones donde el Patrón tenga contratos de prestación de servicios. Dichos cambios no constituirán modificación unilateral de las condiciones de trabajo siempre que se respeten los derechos laborales del trabajador.

	CUARTA. El Trabajador se obliga a cumplir estrictamente las consignas generales y particulares establecidas para cada servicio, incluyendo procedimientos de acceso, control de visitantes, vigilancia perimetral, rondines y reportes.  De conformidad con el Reglamento Interior de Trabajo de la Empresa, las funciones del TRABAJADOR están limitadas estrictamente a la prevención, vigilancia y control de accesos. En caso de emergencias (tales como siniestros, robos en proceso o accidentes), la intervención del TRABAJADOR se limitará de manera enunciativa más no limitativa a: activar los protocolos de seguridad pasiva, dar aviso inmediato a los cuerpos de auxilio públicos (policía, bomberos, ambulancias), reportar a la central de operaciones de El PATRÓN y auxiliar en la evacuación segura del personal. Queda estrictamente prohibido realizar acciones de confrontación o tácticas que pongan en riesgo su integridad física o la de terceros. La omisión injustificada en el cumplimiento de estas consignas será sancionada conforme a la Ley Federal del Trabajo y al Reglamento Interior de Trabajo.

	QUINTA. El PATRÓN pagará al TRABAJADOR, por los servicios prestados de conformidad con este contrato, un salario ordinario de {datos['Salario Semanal']} pesos semanales, el cual se cubrirá los viernes de cada semana. En este importe ya se encuentra incluido el pago correspondiente a los séptimos días (días de descanso semanal) y los días festivos de descanso obligatorio en términos de los artículos 69 y 74 de la Ley Federal del Trabajo. 
De los Bonos de Asistencia y Puntualidad: Adicionalmente al salario ordinario, el PATRÓN otorgará al TRABAJADOR un Bono de Asistencia Semanal por la cantidad de $450.00 pesos (cuatrocientos cincuenta pesos 00/100 M.N.) y un Bono de Puntualidad Semanal por la cantidad de $450.00 pesos (cuatrocientos cincuenta pesos 00/100 M.N.).
El TRABAJADOR queda estrictamente obligado al cumplimiento del 100% de sus asistencias y de sus horarios de entrada durante la semana correspondiente para devengar dichos conceptos. Las partes acuerdan que el nacimiento del derecho a recibir estos bonos está condicionado estrictamente al cumplimiento perfecto de la asistencia y puntualidad; por lo tanto, en caso de que el TRABAJADOR incurra en una sola falta de asistencia (justificada o injustificada) o en un solo retardo durante el periodo semanal, no se generará ni se pagará el bono correspondiente a la falta incurrida (ya sea de asistencia, de puntualidad, o ambos), sin responsabilidad alguna para el PATRÓN.
En caso de que el día de pago sea de descanso obligatorio o festivo bancario, el depósito se efectuará el día hábil inmediato anterior. El pago del salario y de las prestaciones que correspondan se realizará exclusivamente mediante transferencia electrónica de fondos a la cuenta bancaria institucional que EL PATRÓN apertura a nombre del TRABAJADOR, o a la cuenta que este designe por escrito. Los costos de apertura y manejo de cuenta correrán por cuenta de la Empresa. Al importe del salario y bonos se le realizarán las deducciones legales de impuestos y Seguridad Social correspondientes. 
	Por su parte, el TRABAJADOR se obliga, en cualquier caso, a firmar el recibo correspondiente por los pagos efectuados. Si, por alguna razón, éste no firmara el recibo, las partes aceptan que el simple depósito bancario produce efecto liberatorio de pago para el PATRÓN.  

	SEXTA. La duración máxima de la semana laboral será de 45 (cuarenta y cinco) horas, distribuidas de lunes a sábado de cada semana, de conformidad con lo dispuesto por el segundo párrafo del artículo 59 de la Ley Federal del Trabajo. 
Debido a la naturaleza especializada de las actividades de seguridad y vigilancia que presta la Empresa, y para garantizar la continuidad y la cobertura ininterrumpida de los servicios contratados por nuestros clientes, las partes acuerdan expresamente que las jornadas y horarios de trabajo no serán fijos ni permanentes. El TRABAJADOR prestará sus servicios bajo esquemas de turnos rotativos (tales como 8x16 horas, 12x12 horas, 24x24 horas, o los esquemas que operativamente se requieran) , según el rol que le sea comunicado oportunamente por su superior inmediato o supervisor de zona. 
Dentro de la jornada continua, el TRABAJADOR dispondrá de un lapso de 30 (treinta) minutos intermedios para tomar alimentos y reposar, el cual será considerado como tiempo efectivo de trabajo y se adaptará de forma flexible a las necesidades de cada servicio ; durante este periodo, el trabajador podrá utilizar las sillas o asientos destinados para tal efecto de acuerdo con el artículo 132 fracción V de la Ley Federal del Trabajo. El PATRÓN podrá modificar en cualquier tiempo el horario, rol y la rotación de turnos conforme a las necesidades operativas de la Empresa. 
El TRABAJADOR no laborará tiempo extra en su jornada normal, ni durante días de descanso, salvo previa orden expresa y por escrito emitida por el representante del PATRÓN ; orden sin la cual no se reconocerá ni se pagará tiempo extraordinario alguno. 

	SEPTIMA. Las partes convienen en que los días de descanso semanal serán el domingo, sin perjuicio de que el PATRÓN modifique dichos días de descanso semanal cuando las necesidades del servicio así lo requieran. 

	OCTAVA. Cuando el TRABAJADOR tenga más de un año de servicios, disfrutará de doce días de vacaciones anuales en los términos y condiciones que establece la Ley Federal del Trabajo. Adicionalmente, el PATRÓN pagará al TRABAJADOR una prima vacacional del 25% sobre el salario que le corresponda por sus días de vacaciones, en términos de lo dispuesto por el Artículo 80 de la Ley Federal del Trabajo.

	NOVENA. Serán días de descanso obligatorio los que señala el Artículo 74 de la Ley Federal del Trabajo. Atendiendo a la naturaleza especializada de los servicios de seguridad y vigilancia de la Empresa, la cobertura de estos días estará sujeta a los roles operativos cambiantes y requerimientos de los clientes asignados por el supervisor de zona, obligándose el TRABAJADOR a prestar sus servicios si la operación lo requiere, previo pago de las compensaciones legales aplicables.

	DECIMA. El PATRÓN pagará al TRABAJADOR un aguinaldo anual, equivalente a 15 días de salario, en los términos que establece el Artículo 87 de la Ley Federal del Trabajo, mismo que deberá cubrirse a más tardar el 20 de diciembre de cada año.

	DECIMA PRIMERA. EQUIPO Y UNIFORMES. El Patrón proporcionará los uniformes, gafetes, equipo de protección y herramientas necesarias para el desempeño de sus funciones. Los radios, teléfonos, cámaras corporales, dispositivos electrónicos y demás equipos proporcionados por la empresa deberán utilizarse exclusivamente para fines laborales.
Queda prohibido alterar, modificar, cambiar, dañar, prestar o utilizar dichos equipos para fines personales. El Trabajador se obliga a utilizarlos adecuadamente y devolverlos al concluir la relación laboral.

	DECIMA SEGUNDA. OBLIGACIONES DEL TRABAJADOR. Además de las obligaciones previstas en el artículo 134 de la Ley Federal del Trabajo, el TRABAJADOR se obliga estrictamente a cumplir con lo establecido en el Reglamento Interior de Trabajo de la Empresa, comprometiéndose de manera enunciativa más no limitativa a: 
Cumplimiento de Instrucciones y Consignas: Acatar con eficacia, cuidado y esmero las instrucciones de trabajo, órdenes patronales y las consignas específicas o particulares establecidas para el puesto o servicio asignado. 
Puntualidad y Asistencia: Presentarse puntualmente a sus labores respetando los horarios y roles de turnos asignados, registrando personalmente su entrada y salida en los controles (tarjeta, bitácora o lector) que determine el PATRÓN. 
Permanencia y Protocolo de Relevo: Permanecer de forma estrictamente personal en su puesto de vigilancia y control de accesos hasta que se presente físicamente su relevo y se realice la entrega formal de la bitácora y equipo; reconociendo que el abandono del puesto sin autorización expresa del supervisor constituirá una falta grave. 
Estricta Confidencialidad y Reserva: Guardar absoluta reserva y discreción sobre los asuntos de la Empresa y sus clientes. Queda estrictamente prohibido fotografiar, videograbar, reproducir, extraer o difundir por cualquier medio (incluyendo redes sociales o WhatsApp) el contenido de bitácoras, controles de acceso, sistemas de monitoreo o pantallas de CCTV. 
Uso y Portación del Uniforme: Mantener una imagen personal aseada y profesional, portando correctamente el uniforme completo, limpio, fajado y con el gafete de identificación visible durante toda su jornada laboral, como medida de confianza hacia los clientes. 
Inspección y Conservación de Equipo: Revisar, inspeccionar y conservar en buen estado las herramientas, útiles, uniformes y el equipo táctico o de comunicación (radios, fornituras, linternas) proporcionados por el PATRÓN; reportando inmediatamente cualquier desperfecto y absteniéndose de utilizarlos para fines personales. 
Reporte de Incidencias: Informar de manera inmediata a su supervisor de zona o a la central de operaciones sobre cualquier incidente de seguridad, anomalía, siniestro o situación de riesgo detectada en las instalaciones. 
Conducta Profesional: Mantener en todo momento un trato respetuoso, digno y profesional con los clientes, visitantes, proveedores y compañeros de trabajo. 
Exámenes de Control y Confianza: Someterse a los exámenes médicos, psicométricos, de alcoholemia y toxicológicos aleatorios o periódicos que determine la Empresa o las autoridades competentes para prevenir riesgos de trabajo; aceptando que la negativa a realizárselos será causa de rescisión inmediata de la relación de trabajo. 
Medidas Preventivas de Seguridad e Higiene: Observar y acatar rigurosamente todas las medidas preventivas, higiénicas y de seguridad que acuerden las autoridades y las que indique el PATRÓN para salvaguardar su integridad física, la de sus compañeros y la de los bienes resguardados. 

	DECIMA TERCERA. El trabajador deberá elaborar y entregar oportunamente los reportes, bitácoras, formatos de novedades y demás documentos operativos requeridos por la empresa. La falsificación de información o la omisión deliberada de hechos relevantes constituirá falta grave.

	DECIMA CUARTA. El TRABAJADOR se obliga a obedecer estrictamente las normas de trabajo fijadas por el PATRÓN y a respetar la organización jerárquica que la misma tiene establecidas o en el futuro establezca, así como a cumplir con todas las obligaciones que naturalmente deriven de este contrato y de los servicios que debe prestar. 

	DECIMA QUINTA. El TRABAJADOR reconoce que son propiedad exclusiva del PATRÓN y/o sus clientes todos los documentos e información que se le proporcionen con motivo de la relación de trabajo, así como los que el propio PATRÓN prepare o formule en relación o conexión con sus servicios, por lo que se obliga a conservarlos en buen estado y a entregarlos al PATRÓN en el momento en que éste lo requiera o bien al terminar el presente contrato, por el motivo que sea.

	DECIMA SEXTA. El TRABAJADOR se obliga a devolver, a satisfacción del PATRÓN, los instrumentos, equipos y materiales que le fueren proporcionados para el desempeño de sus funciones en el momento en que éste lo requiera o al término del presente acuerdo de voluntades por el motivo que fuere.
Si el EMPLEADO dejare de cumplir con lo establecido en la presente Clausula, quedará sujeto a la responsabilidad civil por los daños o perjuicios que causare al PATRÓN como dueño y propietario de estas herramientas de trabajo, así como las sanciones de carácter penal a que por ello se hiciere acreedor.     

	DECIMA SEPTIMA. CONFIDENCIALIDAD. 
El TRABAJADOR, en cumplimiento a la fracción XIII del artículo 134 de la Ley Federal del Trabajo, se obliga a no divulgar ninguno de los aspectos de los negocios del PATRÓN, información de clientes, vulnerabilidades de clientes, manuales y datos personales a los que tenga acceso con motivo de su trabajo.ni datos personales a terceras personas, verbalmente o por escrito, directa o indirectamente, información alguna sobre los sistemas o actividades de cualquier clase que observe el PATRÓN. La revelación de esta información a terceros será causa de rescisión de la relación laboral.
Si el TRABAJADOR dejare de cumplir con las disposiciones de esta Cláusula, quedará sujeto a la responsabilidad civil por daños y perjuicios que cause al PATRÓN y a las sanciones penales que marca la ley a que se haga acreedor.

	DECIMA OCTAVA. El TRABAJADOR se obliga a cumplir con las disposiciones legales en materia de protección de datos personales y a resguardar adecuadamente cualquier información a la que tenga acceso durante la prestación de sus servicios.

	DECIMA NOVENA. Conforme a lo dispuesto por la Fracción X del Artículo 134 de la Ley Federal del Trabajo, el TRABAJADOR se someterá a los exámenes médicos que ordene el PATRÓN, en la inteligencia de que el facultativo que los practique será designado y retribuido por éste mismo. 

	VIGESIMA. Para todo lo relacionado con riesgos de trabajo y enfermedades o accidentes no profesionales, se estará a lo dispuesto por la Ley del Seguro Social y sus Reglamentos, para lo cual el PATRÓN inscribirá oportunamente al TRABAJADOR ante el Instituto Mexicano del Seguro Social, cubriéndose las cuotas por ambas partes, en los términos que consigna la citada Ley.
En tal virtud, el único documento válido para justificar faltas de asistencia derivadas de incapacidad por enfermedad general o profesional, será el certificado y/o incapacidad que expida el Instituto Mexicano del Seguro Social.

	VIGESIMA PRIMERA. El PATRÓN proporcionará capacitación y adiestramiento al TRABAJADOR, conforme a los planes y programas establecidos, o que se establezcan de acuerdo con las disposiciones de la Ley Federal del Trabajo, comprometiéndose el TRABAJADOR a dedicar el tiempo y esfuerzo necesarios para lograr la mejor eficiencia en dicha capacitación.

	VIGESIMA SEGUNDA. Las partes convienen que, en lo no previsto por el presente contrato, se sujetarán a las disposiciones de la Ley Federal del Trabajo y de la Ley del Seguro Social y sus Reglamentos

	VIGESIMA TERCERA. CAUSAS DE RESCISIÓN:
Serán causas de rescisión sin responsabilidad para el Patrón las previstas en el artículo 47 de la Ley Federal del Trabajo y demás disposiciones aplicables.

	VIGESIMA CUARTA. Salvo autorización expresa del supervisor o cliente, queda restringido el uso de teléfonos celulares, audífonos, tabletas u otros dispositivos electrónicos durante la prestación del servicio cuando ello afecte la vigilancia o seguridad del puesto asignado.

	VIGESIMA QUINTA. REGLAMENTO INTERIOR DE TRABAJO.
El Trabajador manifiesta conocer y aceptar el Reglamento Interior de Trabajo de la empresa, obligándose a cumplirlo en todos sus términos y demás disposiciones internas de la empresa.

	VIGESIMA SEXTA. LEGISLACIÓN APLICABLE.
Para todo lo no previsto en este contrato serán aplicables la Ley Federal del Trabajo, la legislación laboral vigente, la normativa de seguridad social y las disposiciones aplicables a la prestación de servicios de seguridad privada en el Estado de Nuevo León.

	VIGÉSIMA SEPTIMA. Cualquier modificación que se haga al presente contrato deberá constar por escrito y deberá ser firmada de conformidad por ambas partes.

Leído que lo fue íntegramente el presente contrato y enteradas las partes de su contenido y alcancel legal, el TRABAJADOR y el PATRÓN lo ratificaron y firmaron de conformidad, por duplicado, ante dos testigos, en la calle Santa Barbara número 141, Colonia Valle de Santa Isabel, C.P. 67256, Ciudad Benito Juarez, Nuevo León, el _____ de _________________ de 202__, quedando un original en poder del PATRÓN y otro en poder del TRABAJADOR.

EL PATRON
AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE S.A. DE C.V.
C. ABNER VELAZQUEZ MORALES
Representante legal
 | EL TRABAJADOR
{datos['Nombre']}
Por sus propios derechos

TESTIGO
(Nombre)
(Dirección)
 | TESTIGO
(Nombre)
(Dirección)
"""

            for parrafo in texto_contrato.split("\n\n"):
                p_limpio = parrafo.strip()
                if p_limpio:
                    doc.add_paragraph(p_limpio)

            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            registrar_auditoria(
                st.session_state.usuario_actual,
                "CONTRATO",
                f"Generó contrato tipo '{tipo_c}' para {datos['Nombre']}",
            )
            st.success(f"¡Contrato de {tipo_c} generado con éxito!")
            st.download_button(
                "📥 Descargar Documento Word",
                buffer,
                file_name=f"Contrato_{tipo_c.split()[0]}_{datos['Nombre'].replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

# --- 🛡️ PANEL DE ADMINISTRADOR MAESTRO Y GESTIÓN DE USUARIOS ---
elif menu == "🛡️ Panel de Administrador (Usuarios y Auditoría)":
    if st.session_state.rol_actual != "Administrador":
        st.error(
            "⛔ Acceso restringido exclusivamente al Administrador Maestro."
        )
    else:
        st.header("🛡️ Panel de Control y Seguridad del Administrador")
        st.markdown(
            "Gestión centralizada de credenciales de acceso, altas de usuarios y bitácora de auditoría."
        )

        tab_usuarios, tab_credenciales, tab_auditoria = st.tabs(
            [
                "👥 Altas y Gestión de Usuarios",
                "🔑 Cambiar Mis Credenciales de Admin",
                "📋 Bitácora de Auditoría",
            ]
        )

        with tab_usuarios:
            st.subheader("➕ Dar de Alta Nuevo Usuario (Operador / Admin)")

            with st.form("form_nuevo_usuario_maestro"):
                col_u1, col_u2 = st.columns(2)
                with col_u1:
                    nuevo_user = st.text_input("Nombre de Usuario Nuevo")
                    nuevo_rol = st.selectbox(
                        "Rol de Acceso", ["Operador", "Administrador"]
                    )
                with col_u2:
                    nuevo_pass = st.text_input(
                        "Contraseña Temporal", type="password"
                    )

                btn_crear = st.form_submit_button(
                    "💾 Registrar Nuevo Usuario en el Sistema"
                )

                if btn_crear:
                    if nuevo_user and nuevo_pass:
                        lista_u = cargar_usuarios()
                        if any(u["Usuario"] == nuevo_user for u in lista_u):
                            st.error(
                                f"El usuario '{nuevo_user}' ya existe en el sistema."
                            )
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
                                f"¡Usuario '{nuevo_user}' registrado con éxito!"
                            )
                    else:
                        st.error(
                            "Por favor complete el nombre de usuario y la contraseña."
                        )

            st.markdown("---")
            st.subheader("📋 Usuarios Activos en la Aplicación")
            usuarios_actuales = cargar_usuarios()
            df_usuarios = pd.DataFrame(usuarios_actuales)[["Usuario", "Rol"]]
            st.dataframe(df_usuarios, use_container_width=True)

            st.markdown("### 🗑️ Revocar Acceso a Usuario")
            nombres_usuarios_del = [
                u["Usuario"]
                for u in usuarios_actuales
                if u["Usuario"] != st.session_state.usuario_actual
            ]
            if nombres_usuarios_del:
                user_a_borrar = st.selectbox(
                    "Seleccione usuario a eliminar", nombres_usuarios_del
                )
                if st.button("❌ Eliminar Acceso de este Usuario"):
                    usuarios_actuales = [
                        u
                        for u in usuarios_actuales
                        if u["Usuario"] != user_a_borrar
                    ]
                    guardar_usuarios(usuarios_actuales)
                    registrar_auditoria(
                        st.session_state.usuario_actual,
                        "BAJA USUARIO",
                        f"Eliminó al usuario '{user_a_borrar}'",
                    )
                    st.warning(
                        f"Acceso revocado para el usuario '{user_a_borrar}'."
                    )
                    st.rerun()
            else:
                st.info(
                    "No hay usuarios adicionales para eliminar (tu usuario actual está protegido)."
                )

        with tab_credenciales:
            st.subheader(
                "🔑 Actualizar Mi Nombre de Usuario y Contraseña de Administrador"
            )

            with st.form("form_cambiar_admin"):
                pass_actual = st.text_input(
                    "Contraseña Actual", type="password"
                )
                nuevo_admin_user = st.text_input(
                    "Nuevo Nombre de Administrador",
                    value=st.session_state.usuario_actual,
                )
                nuevo_admin_pass = st.text_input(
                    "Nueva Contraseña", type="password"
                )
                confirma_admin_pass = st.text_input(
                    "Confirmar Nueva Contraseña", type="password"
                )

                btn_actualizar_admin = st.form_submit_button(
                    "🔐 Guardar Cambios de Credenciales"
                )

                if btn_actualizar_admin:
                    lista_u = cargar_usuarios()
                    admin_actual = next(
                        (
                            u
                            for u in lista_u
                            if u["Usuario"]
                            == st.session_state.usuario_actual
                        ),
                        None,
                    )

                    if (
                        admin_actual
                        and str(admin_actual["Password"]) == pass_actual
                    ):
                        if nuevo_admin_pass == confirma_admin_pass:
                            if len(nuevo_admin_pass) >= 4:
                                for u in lista_u:
                                    if (
                                        u["Usuario"]
                                        == st.session_state.usuario_actual
                                    ):
                                        u["Usuario"] = nuevo_admin_user
                                        u["Password"] = nuevo_admin_pass

                                guardar_usuarios(lista_u)
                                registrar_auditoria(
                                    st.session_state.usuario_actual,
                                    "CAMBIO CREDENCIALES",
                                    "El administrador actualizó su usuario/contraseña",
                                )
                                st.session_state.usuario_actual = (
                                    nuevo_admin_user
                                )
                                st.success(
                                    "¡Credenciales actualizadas con éxito! Vuelve a iniciar sesión si es necesario."
                                )
                            else:
                                st.error(
                                    "La contraseña debe tener al menos 4 caracteres."
                                )
                        else:
                            st.error(
                                "Las nuevas contraseñas no coinciden. Verifícalas."
                            )
                    else:
                        st.error(
                            "La contraseña actual es incorrecta. No se pudieron aplicar los cambios."
                        )

        with tab_auditoria:
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
