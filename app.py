import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import os

# Configuración de la página
st.set_page_config(
    page_title="AVM - Sistema de Recursos Humanos", page_icon="🛡️", layout="wide"
)

# Aplicar diseño corporativo en Negro y Dorado mediante CSS
st.markdown(
    """
    <style>
        /* Fondo general de la aplicación */
        .stApp {
            background-color: #0e0e0e;
            color: #f3f3f3;
        }
        
        /* Barra lateral */
        [data-testid="stSidebar"] {
            background-color: #161616;
            border-right: 1px solid #d4af37;
        }
        
        /* Títulos y textos principales */
        h1, h2, h3, h4, h5, h6, span, label {
            color: #f3f3f3 !important;
        }
        
        /* Acento dorado para títulos principales */
        h1 {
            color: #d4af37 !important;
            border-bottom: 2px solid #d4af37;
            padding-bottom: 10px;
        }
        
        /* Botones personalizados con bordes y acentos dorados */
        .stButton>button {
            background-color: #d4af37;
            color: #0e0e0e;
            font-weight: bold;
            border: none;
            border-radius: 4px;
            padding: 0.5rem 1rem;
        }
        
        .stButton>button:hover {
            background-color: #f3e5ab;
            color: #000000;
        }
        
        /* Tarjetas de formularios y contenedores */
        div.stForm {
            background-color: #1a1a1a;
            padding: 20px;
            border-radius: 8px;
            border: 1px solid #333333;
        }
        
        /* Campos de texto y selectores */
        input, select, textarea {
            background-color: #222222 !important;
            color: #ffffff !important;
            border: 1px solid #444444 !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# Encabezado con Logotipo (Buscando el archivo local 'logo.png' en tu repo)
# CORREGIDO: Se usa st.columns en plural
col_logo, col_titulo = st.columns([1, 4])

with col_logo:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=130)
    else:
        st.markdown("🛡️ **[Sube tu 'logo.png' al repositorio]**")

with col_titulo:
    st.title("AVM Grupo Integral de Seguridad Privada del Norte")
    st.subheader("Sistema de Gestión de Recursos Humanos y Contratos")

# Menú lateral
menu = st.sidebar.selectbox(
    "Menú de Navegación",
    [
        "Registro de Personal",
        "Generar Contrato Sujeto a Prueba",
        "Generar Contrato Tiempo Indeterminado",
    ],
)

# Base de datos simulada en memoria de la sesión
if "empleados" not in st.session_state:
    st.session_state.empleados = []

if menu == "Registro de Personal":
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
                    "domicilio": domicilio,
                    "puesto": puesto,
                    "salario_semanal": salario_semanal,
                }
                st.session_state.empleados.append(nuevo_emp)
                st.success(
                    f"¡Guardia {nombre} registrado correctamente en el sistema!"
                )
            else:
                st.error("Por favor ingresa al menos el Nombre y la CURP.")

    if len(st.session_state.empleados) > 0:
        st.subheader("📋 Personal Registrado Recientemente")
        df = pd.DataFrame(st.session_state.empleados)
        st.dataframe(df)


elif menu == "Generar Contrato Sujeto a Prueba":
    st.header("📄 Generador de Contrato - Sujeto a Prueba (30 Días)")

    if not st.session_state.empleados:
        st.warning(
            "⚠️ Primero registra un empleado en la sección 'Registro de Personal'."
        )
    else:
        nombres_empleados = [e["nombre"] for e in st.session_state.empleados]
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
a) Ser una persona moral, debidamente constituida conforme a las leyes de la República Mexicana, según consta en la escritura pública número 6,948, pasada ante la fe del Notario Público número 127, con domicilio ubicado en Calle Santa Bárbara número 141, C. Asturias, Colonia Valle de Santa Isabel, C.P. 67256, Ciudad Benito Juárez, Nuevo León, Registro Federal de Contribuyentes AGI260413CK4 y tener como objeto social, entre otros, la prestación de servicios de seguridad privada, consistentes en la vigilancia, protección y resguardo de bienes muebles e inmuebles, así como de establecimientos comerciales, industriales, habitacionales y de servicios.
b) Que requiere de personal capacitado para ocupar el puesto de {datos['puesto']}.

II. Declara el TRABAJADOR:
a) Ser una persona física, de nacionalidad {datos['nacionalidad']}, sexo {datos['sexo']}, fecha de nacimiento {datos['fecha_nacimiento']}, estado civil {datos['estado_civil']}, CURP {datos['curp']} y RFC {datos['rfc']}, con domicilio en {datos['domicilio']}.

C L Á U S U L A S:
PRIMERA. El presente contrato se celebra por TIEMPO INDETERMINADO, quedando sujeto "EL TRABAJADOR" a un PERIODO DE PRUEBA DE 1 MES (30 DÍAS).
QUINTA. El PATRÓN pagará al TRABAJADOR un salario ordinario de {datos['salario_semanal']} pesos semanales, más Bonos de Asistencia ($450.00) y Puntualidad ($450.00).
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


elif menu == "Generar Contrato Tiempo Indeterminado":
    st.header("📄 Generador de Contrato - Tiempo Indeterminado")

    if not st.session_state.empleados:
        st.warning(
            "⚠️ Primero registra un empleado en la sección 'Registro de Personal'."
        )
    else:
        nombres_empleados = [e["nombre"] for e in st.session_state.empleados]
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

            texto_indet = f"""CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO INDETERMINADO, QUE CELEBRAN, POR UNA PARTE, AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE, SOCIEDAD ANONIMA DE CAPITAL VARIABLE, REPRESENTADA POR EL C. ABNER VELAZQUEZ MORALES (EL "PATRÓN"), Y POR LA OTRA PARTE, {datos['nombre']} (EL “TRABAJADOR”), CONFORME A LAS SIGUIENTES CLÁUSULAS:

PRIMERA. El TRABAJADOR se obliga a prestar sus servicios con el puesto de {datos['puesto']}.
QUINTO. Salario semanal de {datos['salario_semanal']}, cubriendo bonos de asistencia y puntualidad condicionados al 100% de asistencia.
"""
            for parrafo in texto_indet.split("\n\n"):
                if parrafo.strip():
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
