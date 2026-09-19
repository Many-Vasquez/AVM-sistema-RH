import streamlit as st
import pandas as pd
from datetime import datetime
from PIL import Image
import requests
from io import BytesIO
import re
from docx import Document

# --- CONFIGURACIÓN DE LA PÁGINA Y LOGO ---
# Reemplaza esta URL con el enlace directo de tu logo si deseas cambiarlo
URL_LOGO = "https://drive.google.com/file/d/1ev83sbeISkt451XlJ7jApeNj786lp1ao/view?usp=drive_link" 

@st.cache_resource
def load_image_from_url(url):
    try:
        response = requests.get(url)
        return Image.open(BytesIO(response.content))
    except:
        return None

logo = load_image_from_url(URL_LOGO)
st.set_page_config(page_title="AVM Seguridad - Gestión de Personal", page_icon=logo if logo else "🛡️", layout="wide")

# --- FUNCIONES DE VALIDACIÓN OFICIAL (EXPRESIONES REGULARES MÉXICO) ---
def validar_curp(curp):
    patron = r"^[A-Z]{4}[0-9]{6}[HM][A-Z]{5}[0-9A-Z]{2}$"
    return bool(re.match(patron, curp))

def validar_rfc(rfc):
    # Soporta RFC de persona física (13 caracteres)
    patron = r"^[A-Z&Ñ]{4}[0-9]{6}[A-Z0-9]{3}$"
    return bool(re.match(patron, rfc))

def validar_nss(nss):
    patron = r"^[0-9]{11}$"
    return bool(re.match(patron, nss))

# --- INICIALIZACIÓN DE LA BASE DE DATOS EN MEMORIA ---
if 'empleados' not in st.session_state:
    st.session_state['empleados'] = pd.DataFrame(columns=[
        "ID", "Nombre Completo", "CURP", "RFC", "NSS", "Dirección", 
        "Teléfono", "Estado Civil", "Puesto", "Salario Diario", "Fecha de Alta"
    ])

# --- INTERFAZ VISUAL PRINCIPAL ---
if logo:
    col_logo, col_titulo = st.columns([1, 6])
    with col_logo:
        st.image(logo, width=120)
    with col_titulo:
        st.title("AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE")
        st.subheader("Sistema Integral de Recursos Humanos y Contratos")
else:
    st.title("🛡️ AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE")
    st.subheader("Sistema Integral de Recursos Humanos y Contratos")

# Menú lateral de navegación
menu = st.sidebar.selectbox("Menú Principal", ["Dashboard / Empleados", "Nuevo Registro & Contrato"])

# =========================================================================
# SECCIÓN 1: DASHBOARD Y LISTADO
# =========================================================================
if menu == "Dashboard / Empleados":
    st.markdown("### 📋 Listado de Personal Activo")
    df = st.session_state['empleados']
    
    if not df.empty:
        st.dataframe(df, use_container_width=True)
        st.metric("Total de Elementos Registrados", len(df))
    else:
        st.info("No hay elementos registrados todavía. Utiliza la sección 'Nuevo Registro & Contrato' para dar de alta al primero.")

# =========================================================================
# SECCIÓN 2: NUEVO REGISTRO Y VALIDACIÓN
# =========================================================================
elif menu == "Nuevo Registro & Contrato":
    st.markdown("### ✍️ Alta de Nuevo Elemento y Datos Contractuales")
    
    with st.form("form_empleado"):
        col1, col2 = st.columns(2)
        
        with col1:
            # Forzamos texto en mayúsculas automáticamente mediante .upper()
            nombre = st.text_input("Nombre Completo (Empezando por Apellidos)").upper()
            curp = st.text_input("CURP (18 caracteres)").upper()
            rfc = st.text_input("RFC con Homoclave (13 caracteres)").upper()
            nss = st.text_input("NSS (11 dígitos)").upper()
            direccion = st.text_input("Dirección Completa (Calle, Núm, Colonia, Municipio)").upper()
            
        with col2:
            telefono = st.text_input("Teléfono de Contacto (10 dígitos)")
            estado_civil = st.selectbox("Estado Civil", ["SOLTERO(A)", "CASADO(A)", "UNIÓN LIBRE", "DIVORCIADO(A)", "VIUDO(A)"])
            puesto = st.selectbox("Puesto", ["GUARDIA INTRAMUROS", "SUPERVISOR DE OPERACIONES", "CUSTODIO", "JEFE DE TURNO"])
            salario_diario = st.number_input("Salario Diario (MXN)", min_value=250.0, value=300.0, step=10.0)
            fecha_alta = st.date_input("Fecha de Ingreso", datetime.now())
            
        tipo_contrato = st.selectbox("Tipo de Contrato a Generar", ["SUJETO A PRUEBA", "TIEMPO DETERMINADO"])
        
        submitted = st.form_submit_button("Validar, Guardar y Preparar Contrato")
        
        if submitted:
            # Validaciones estrictas
            val_c = validar_curp(curp)
            val_r = validar_rfc(rfc)
            val_n = validar_nss(nss)
            
            if not (val_c and val_r and val_n):
                st.error("⚠️ Error en las validaciones oficiales:")
                if not val_c: st.markdown("- **CURP inválida**: Revisa la estructura de 18 caracteres.")
                if not val_r: st.markdown("- **RFC inválido**: Revisa la estructura de 13 caracteres con homoclave.")
                if not val_n: st.markdown("- **NSS inválido**: Debe contener exactamente 11 dígitos numéricos.")
            elif not nombre or not direccion or not telefono:
                st.warning("⚠️ Todos los campos de texto y dirección son obligatorios.")
            else:
                # Simulación de cálculo SBC (Factor mínimo primer año)
                factor_integracion = 1.0493 
                sbc = salario_diario * factor_integracion
                
                nuevo_registro = pd.DataFrame({
                    "ID": [len(st.session_state['empleados']) + 1],
                    "Nombre Completo": [nombre],
                    "CURP": [curp],
                    "RFC": [rfc],
                    "NSS": [nss],
                    "Dirección": [direccion],
                    "Teléfono": [telefono],
                    "Estado Civil": [estado_civil],
                    "Puesto": [puesto],
                    "Salario Diario": [f"${salario_diario:,.2f}"],
                    "Fecha de Alta": [str(fecha_alta)]
                })
                
                st.session_state['empleados'] = pd.concat([st.session_state['empleados'], nuevo_registro], ignore_index=True)
                st.success("✅ ¡Elemento validado y registrado exitosamente en el sistema!")
                
                # Guardamos temporalmente en sesión para la descarga del contrato
                st.session_state['ultimo_registrado'] = {
                    "nombre": nombre, "curp": curp, "rfc": rfc, "nss": nss,
                    "direccion": direccion, "telefono": telefono, "estado_civil": estado_civil,
                    "puesto": puesto, "salario": f"${salario_diario:,.2f}", "fecha": str(fecha_alta),
                    "tipo_contrato": tipo_contrato
                }

    # --- GENERADOR DE CONTRATO WORD AUTOMATIZADO ---
    if 'ultimo_registrado' in st.session_state:
        st.markdown("---")
        st.markdown("### 📄 Generación de Contrato Laboral")
        elem = st.session_state['ultimo_registrado']
        
        st.info(f"Listo para generar contrato por **{elem['tipo_contrato']}** para: **{elem['nombre']}**")
        
        if st.button("📥 Descargar Contrato en Formato Word (.docx)"):
            doc = Document()
            doc.add_heading("AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE S.A. DE C.V.", level=1)
            doc.add_heading(f"CONTRATO INDIVIDUAL DE TRABAJO POR {elem['tipo_contrato']}", level=2)
            
            doc.add_paragraph(f"Fecha de elaboración: {datetime.now().strftime('%d/%m/%Y')}\n")
            doc.add_paragraph(f"EL PATRÓN: AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE S.A. DE C.V., representada legalmente.")
            doc.add_paragraph(f"EL TRABAJADOR:\n"
                              f"- Nombre: {elem['nombre']}\n"
                              f"- CURP: {elem['curp']}\n"
                              f"- RFC: {elem['rfc']}\n"
                              f"- NSS: {elem['nss']}\n"
                              f"- Dirección: {elem['direccion']}\n"
                              f"- Teléfono: {elem['telefono']}\n"
                              f"- Estado Civil: {elem['estado_civil']}")
            
            doc.add_heading("CLÁUSULAS PRINCIPALES:", level=3)
            doc.add_paragraph(f"PRIMERA.— El trabajador prestará sus servicios desempeñando el puesto de {elem['puesto']}.")
            doc.add_paragraph(f"SEGUNDA.— El presente contrato se celebra bajo la modalidad de {elem['tipo_contrato']}, de conformidad con la Ley Federal del Trabajo.")
            doc.add_paragraph(f"TERCERA.— Se pagará al trabajador un salario diario de {elem['salario']}, cubierto en moneda de curso legal.")
            
            doc.add_paragraph("\n\n\n____________________________________             ____________________________________")
            doc.add_paragraph("       POR LA EMPRESA (AVM GRUPO)                      EL TRABAJADOR")
            
            # Guardar archivo temporal en memoria
            file_path = f"Contrato_{elem['nombre'].replace(' ', '_')}.docx"
            doc.save(file_path)
            
            with open(file_path, "rb") as f:
                st.download_button(
                    label="💾 Hacer clic aquí para descargar el documento Word",
                    data=f,
                    file_name=file_path,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
