import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import os

# Configuración de la página
st.set_page_config(
    page_title="AVM - Sistema de Recursos Humanos", page_icon="🛡️", layout="wide"
)

# Aplicar diseño corporativo en Negro y Dorado con visibilidad corregida para la flecha de la barra lateral
st.markdown(
    """
    <style>
        /* Fondo general de la aplicación */
        .stApp {
            background-color: #0e0e0e;
            color: #f3f3f3;
        }
        
        /* Barra lateral y botón de despliegue */
        [data-testid="stSidebar"] {
            background-color: #161616;
            border-right: 1px solid #d4af37;
        }
        
        /* Hacer visible la flecha de la barra lateral contra el fondo oscuro */
        button[kind="header"] {
            color: #d4af37 !important;
            background-color: #1a1a1a !important;
            border: 1px solid #d4af37 !important;
        }
        
        [data-testid="collapsedControl"] {
            color: #d4af37 !important;
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

# Encabezado con Logotipo
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
a) Ser una persona moral, debidamente constituida conforme a las leyes de la República Mexicana, según consta en la escritura pública número 6,948, pasada ante la fe del Notario Público número 127, con domicilio ubicado en Calle Santa Bárbara número 141, C. Asturias, Colonia Valle de Santa Isabel, C.P. 67256, Ciudad Benito Juárez, Nuevo León, Registro Federal de Contribuyentes AGI260413CK4 y tener como objeto social, entre otros, la prestación de servicios de seguridad privada, consistentes en la vigilancia, protección y resguardo de bienes muebles e inmuebles, así como de establecimientos comerciales, industriales, habitacionales y de servicios, mediante la utilización de recursos humanos debidamente capacitados.
b) Que, para dar cumplimiento a su objeto social, requiere de personal capacitado y con experiencia para ocupar el puesto de {datos['puesto']} para que realice las actividades consistentes, de manera enunciativa mas no limitativa, en: Vigilancia, protección y resguardo de bienes muebles e inmuebles, seguridad intramuros incluyendo el control de accesos y salidas, registro de personas y vehículos, realización de rondines y supervisión interna de instalaciones privadas.

II. Declara el TRABAJADOR:
a) Ser una persona física, de nacionalidad {datos['nacionalidad']}, de sexo {datos['sexo']}, con fecha de nacimiento el {datos['fecha_nacimiento']}, estado civil {datos['estado_civil']}, Clave Única de Registro de Población {datos['curp']} y Registro Federal de Contribuyentes {datos['rfc']}, con domicilio en {datos['domicilio']}.
b) Que cuenta con los conocimientos, habilidades y experiencia necesarios para prestar al PATRÓN los servicios del puesto encomendado.
c) Que está de acuerdo en prestar los servicios descritos en el presente contrato, sujeto a un periodo de prueba de 30 (treinta) días.

III. Declaran ambas partes:
a) Que cuentan con las facultades suficientes para la celebración del presente contrato y obligarse a los términos de este, reconociéndose mutuamente la personalidad con la que comparecen.

C L Á U S U L A S:

PRIMERA. El presente contrato se celebra por TIEMPO INDETERMINADO, quedando sujeto "EL TRABAJADOR" a un PERIODO DE PRUEBA DE 1 MES (30 DÍAS) contados a partir de la fecha de firma del presente contrato, con fundamento en el párrafo segundo del Artículo 39-A de la Ley Federal del Trabajo, toda vez que el puesto a desempeñar requiere de labores técnicas, operativas y/o conocimientos especializados en materia de seguridad privada, prevención de riesgos y manejo de equipos. Durante dicho periodo de prueba, "EL PATRÓN" evaluará si "EL TRABAJADOR" cumple con los requisitos, conocimientos y aptitudes necesarios para el puesto. De no acreditarlos a satisfacción de "EL PATRÓN" mediante la evaluación correspondiente, podrá dar por terminada la relación de trabajo en cualquier momento, sin responsabilidad alguna para la Empresa y sin obligación de pagar indemnización constitucional, procediendo únicamente al pago del finiquito proporcional.

SEGUNDA. Se hace constar que el PATRÓN celebra el presente contrato fundado en las declaraciones del TRABAJADOR en el sentido de que cuenta con los requisitos y conocimientos necesarios para desempeñar adecuadamente las actividades inherentes al cargo. Al término del periodo de prueba, de no acreditar el TRABAJADOR que satisface los requisitos y conocimientos necesarios, se dará por terminada la relación de trabajo sin responsabilidad para el PATRÓN.

TERCERA. El TRABAJADOR prestará sus servicios en el domicilio del PATRÓN o en cualquier otro domicilio en el que se ubiquen las oficinas, centros de trabajo, clientes o instalaciones donde el Patrón tenga contratos de prestación de servicios de seguridad privada vigentes.

CUARTA. El Trabajador se obliga a cumplir estrictamente las consignas generales y particulares establecidas para cada servicio, procedimientos de acceso, control de visitantes, vigilancia perimetral, rondines y reportes. En caso de emergencias, su actuación se limitará estrictamente a activar los protocolos de seguridad pasiva, dar aviso inmediato a los cuerpos de auxilio públicos y a la central de operaciones de El PATRÓN, quedando prohibido realizar acciones de confrontación o tácticas que pongan en riesgo su integridad física o la de terceros.

QUINTA. El PATRÓN pagará al TRABAJADOR un salario ordinario de {datos['salario_semanal']} pesos semanales, el cual se cubrirá los viernes de cada semana. En este importe ya se encuentra incluido el pago correspondiente a los séptimos días y los días festivos de descanso obligatorio. Adicionalmente, el PATRÓN otorgará al TRABAJADOR un Bono de Asistencia Semanal de $450.00 pesos y un Bono de Puntualidad Semanal de $450.00 pesos, condicionados al cumplimiento perfecto del 100% de asistencias y puntualidad. El pago se realizará mediante transferencia electrónica o depósito bancario.

SEXTA. La duración máxima de la semana laboral será de 45 (cuarenta y cinco) horas, distribuidas de lunes a sábado en turnos rotativos (8x16, 12x12, 24x24 horas, etc.). Dentro de la jornada continua, dispondrá de 30 minutos intermedios para alimentos y reposo.

SEPTIMA. Los días de descanso semanal serán el domingo, sin perjuicio de que el PATRÓN modifique dichos días cuando las necesidades del servicio operativo así lo requieran.

OCTAVA. Cuando el TRABAJADOR tenga más de un año de servicios, disfrutará de doce días de vacaciones anuales y una prima vacacional del 25%.

NOVENA. Serán días de descanso obligatorio los que señala el Artículo 74 de la Ley Federal del Trabajo, cubriéndose mediante roles operativos según los requerimientos de los clientes.

DECIMA. El PATRÓN pagará al TRABAJADOR un aguinaldo anual equivalente a 15 días de salario, a más tardar el 20 de diciembre de cada año.

DECIMA PRIMERA. EQUIPO Y UNIFORMES. El Patrón proporcionará los uniformes, gafetes, equipo de protección y herramientas necesarias, los cuales deberán utilizarse exclusivamente para fines laborales y devolverse al concluir la relación.

DECIMA SEGUNDA. OBLIGACIONES DEL TRABAJADOR. Cumplir estrictamente con el Reglamento Interior de Trabajo, mantener puntualidad, asistencia, permanencia en puesto hasta el relevo, estricta confidencialidad, portación correcta del uniforme y someterse a exámenes médicos y toxicológicos.

DECIMA TERCERA a VIGESIMA SEPTIMA. Las partes se obligan a cumplir con las disposiciones de confidencialidad, protección de datos, seguridad social ante el IMSS, capacitación, causas de rescisión conforme al artículo 47 de la Ley Federal del Trabajo, y legislación aplicable en el Estado de Nuevo León.
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

            texto_indet = f"""CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO INDETERMINADO, QUE CELEBRAN, POR UNA PARTE, AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE, SOCIEDAD ANONIMA DE CAPITAL VARIABLE, REPRESENTADA EN ESTE ACTO POR EL C. ABNER VELAZQUEZ MORALES (EN LO SUCESIVO, EL "PATRÓN"), Y POR LA OTRA PARTE, POR SU PROPIO DERECHO, {datos['nombre']} (EN LO SUCESIVO, EL “TRABAJADOR”), DE CONFORMIDAD CON LOS ARTÍCULOS 20, 21, 24, 25, 35, 132, 134 Y DEMÁS RELATIVOS Y APLICABLES DE LA LEY FEDERAL DEL TRABAJO, AL TENOR DE LAS SIGUIENTES DECLARACIONES Y CLÁUSULAS:

D E C L A R A C I O N E S:

I. Declara el PATRÓN:
a) Ser una persona moral, debidamente constituida conforme a las leyes de la República Mexicana, según consta en la escritura pública número 6,948, pasada ante la fe del Notario Público número 127, con domicilio ubicado en Calle Santa Bárbara número 141, C. Asturias, Colonia Valle de Santa Isabel, C.P. 67256, Ciudad Benito Juárez, Nuevo León, Registro Federal de Contribuyentes AGI260413CK4 y tener como objeto social, entre otros, la prestación de servicios de seguridad privada, consistentes en la vigilancia, protección y resguardo de bienes muebles e inmuebles, así como de establecimientos comerciales, industriales, habitacionales y de servicios, mediante la utilización de recursos humanos debidamente capacitados.
b) Que, para dar cumplimiento a su objeto social, requiere de personal capacitado y con experiencia para ocupar el puesto de {datos['puesto']} para que realice las actividades consistentes, de manera enunciativa mas no limitativa, en: Vigilancia, protección y resguardo de bienes muebles e inmuebles, seguridad intramuros incluyendo el control de accesos y salidas, registro de personas y vehículos, realización de rondines y supervisión interna de instalaciones privadas.

II. Declara el TRABAJADOR:
a) Ser una persona física, de nacionalidad {datos['nacionalidad']}, de sexo {datos['sexo']}, con fecha de nacimiento el {datos['fecha_nacimiento']}, estado civil {datos['estado_civil']}, Clave Única de Registro de Población {datos['curp']} y Registro Federal de Contribuyentes {datos['rfc']}, con domicilio en {datos['domicilio']}.
b) Que cuenta con los conocimientos, habilidades y experiencia necesarios para prestar al PATRÓN los servicios del puesto encomendado.
c) Que está de acuerdo en prestar los servicios descritos en el presente contrato por tiempo indeterminado.

III. Declaran ambas partes:
a) Que cuentan con las facultades suficientes para la celebración del presente contrato y obligarse a los términos del mismo, reconociéndose mutuamente la personalidad con la que comparecen.

C L Á U S U L A S:

PRIMERA. El TRABAJADOR se obliga a prestar, bajo la dirección, dependencia y subordinación del PATRÓN, los servicios personales subordinados consistentes en las actividades del puesto de {datos['puesto']}. Las partes están de acuerdo en que los servicios mencionados se estipulan de manera enunciativa y no limitativa.

SEGUNDA. Se hace constar que el PATRÓN celebra el presente contrato fundado en las declaraciones del TRABAJADOR en el sentido de que cuenta con los requisitos y conocimientos necesarios para desempeñar adecuadamente las actividades inherentes al cargo para el que se le contrata.

TERCERA. El TRABAJADOR prestará sus servicios en el domicilio del PATRÓN o en cualquier otro domicilio en el que se ubiquen las oficinas, centros de trabajo, clientes o instalaciones donde el Patrón tenga contratos de prestación de servicios de seguridad privada vigentes.

CUARTA. El Trabajador se obliga a cumplir estrictamente las consignas generales y particulares establecidas para cada servicio, procedimientos de acceso, control de visitantes, vigilancia perimetral y rondines. En caso de emergencias, su actuación se limitará a activar los protocolos de seguridad pasiva y dar aviso a los cuerpos de auxilio y central de operaciones.

QUINTA. El PATRÓN pagará al TRABAJADOR un salario ordinario de {datos['salario_semanal']} pesos semanales, cubriéndose los viernes de cada semana. En este importe ya se incluye el pago correspondiente a los séptimos días y días festivos. Adicionalmente, el PATRÓN otorgará un Bono de Asistencia Semanal de $450.00 pesos y un Bono de Puntualidad Semanal de $450.00 pesos, condicionados al cumplimiento del 100% de asistencia y puntualidad.

SEXTA. La duración máxima de la semana laboral será de 45 (cuarenta y cinco) horas en turnos rotativos (8x16, 12x12, 24x24 horas, etc.). Contará con 30 minutos intermedios para alimentos y reposo.

SEPTIMA. Los días de descanso semanal serán el domingo, sin perjuicio de las modificaciones por necesidades operativas del servicio.

OCTAVA. Cuando el TRABAJADOR tenga más de un año de servicios, disfrutará de doce días de vacaciones anuales y una prima vacacional del 25%.

NOVENA. Serán días de descanso obligatorio los señalados en el Artículo 74 de la Ley Federal del Trabajo.

DECIMA. El PATRÓN pagará al TRABAJADOR un aguinaldo anual equivalente a 15 días de salario, a más tardar el 20 de diciembre de cada año.

DECIMA PRIMERA a VIGESIMA SEPTIMA. Se aplicarán los términos relativos a entrega obligatoria de equipo y uniformes, estricta confidencialidad, protección de datos personales, exámenes médicos y toxicológicos, inscripción ante el IMSS, capacitación, causas de rescisión conforme al artículo 47 de la Ley Federal del Trabajo y normatividad vigente en Nuevo León.
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
