import os
from io import BytesIO
import pandas as pd
from docx import Document
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="AVM - Sistema de Recursos Humanos y Comercial",
    page_icon="🛡️",
    layout="wide",
)

# Aplicar diseño corporativo en Negro y Dorado
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

# Encabezado con Logotipo
col_logo, col_titulo = st.columns([1, 4])

with col_logo:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=130)
    else:
        st.markdown("🛡️ **[Sube tu 'logo.png' al repositorio]**")

with col_titulo:
    st.title("AVM Grupo Integral de Seguridad Privada del Norte")
    st.subheader(
        "Sistema de Gestión de Recursos Humanos y Propuestas Comerciales"
    )

# Menú lateral
menu = st.sidebar.selectbox(
    "Menú de Navegación",
    [
        "Registro de Personal",
        "Generar Contrato Sujeto a Prueba",
        "Generar Contrato Tiempo Indeterminado",
        "Generador de Cotizaciones",
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

QUINTA. El PATRÓN pagará al TRABAJADOR un salario ordinario de {datos['salario_semanal']} pesos semanales, cubriéndose los viernes de cada semana (incluye séptimos días y días festivos). Adicionalmente, el PATRÓN otorgará un Bono de Asistencia Semanal de $450.00 pesos y un Bono de Puntualidad Semanal de $450.00 pesos, condicionados al cumplimiento perfecto del 100% de asistencias y puntualidad.

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



__________________________________



EL TRABAJADOR
{datos['nombre']}



__________________________________
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

QUINTA. El PATRÓN pagará al TRABAJADOR un salario ordinario de {datos['salario_semanal']} pesos semanales, cubriéndose los viernes de cada semana. En este importe ya se incluye el pago correspondiente a los séptimos días y días festivos. Adicionalmente, el PATRÓN otorgará un Bono de Asistencia Semanal de $450.00 pesos y un Bono de Puntualidad Semanal de $450.00 pesos, condicionados al cumplimiento del 100% de asistencia y puntualidad.

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



__________________________________



EL TRABAJADOR
{datos['nombre']}



__________________________________
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


elif menu == "Generador de Cotizaciones":
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

        submitted_cot = st.form_submit_button(
            "📥 Descargar Propuesta Económica (Word)"
        )

        if submitted_cot:
            subtotal = cantidad_guardias * precio_unitario
            iva = subtotal * 0.16
            total = subtotal + iva

            doc_cot = Document()
            doc_cot.add_heading(
                "AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE, SA DE CV",
                level=1,
            )
            doc_cot.add_paragraph(
                "SANTA BARBARA NUMERO 141, COLONIA VALLE DE SANTA ISABEL, C.P. 67256, CIUDAD BENITO JUAREZ, NUEVO LEON\nPROPUESTA ECONOMICA DE SERVICIOS"
            )

            doc_cot.add_paragraph(
                f"FECHA: {fecha_cot}\nEMPRESA: {empresa_cliente}\nCONTACTO: {contacto_cliente}"
            )

            doc_cot.add_heading("ANÁLISIS DE SITUACIÓN", level=2)
            doc_cot.add_paragraph(
                "Tras evaluar las necesidades de seguridad de su instalación, nuestra firma propone un esquema de Seguridad Proactiva. A diferencia de la vigilancia convencional, nuestro servicio se basa en la disuasión avanzada y la capacidad de respuesta inmediata bajo los más altos estándares de cumplimiento legal."
            )

            doc_cot.add_heading("TABLA DE COTIZACIÓN", level=2)
            doc_cot.add_paragraph(
                f"Cantidad: {cantidad_guardias} | Categoría: Guardias Intramuro o extramuro (Control de Accesos y Caseta) | Descripción: Control estricto de acceso peatonal y vehicular (empleados, contratistas, proveedores y transporte pesado). Turno de 12 horas | Precio Unitario: ${precio_unitario:,.2f} | Total Mensual: ${subtotal:,.2f}"
            )
            doc_cot.add_paragraph(
                f"Subtotal: ${subtotal:,.2f}\nIVA (16%): ${iva:,.2f}\nTOTAL: ${total:,.2f}[cite: 10]"
            )

            doc_cot.add_heading("TÉRMINOS Y CONDICIONES COMERCIALES", level=2)
            doc_cot.add_heading(
                "1. Responsabilidad Civil y Patronal", level=3
            )
            doc_cot.add_paragraph(
                "Nuestra firma asume la totalidad de las obligaciones derivadas de las leyes laborales, de seguridad social (IMSS, INFONAVIT) y fiscales vigentes. El cliente queda exento de cualquier responsabilidad solidaria, ya que todo el personal operativo depende directamente de nuestra razón social[cite: 10]."
            )

            doc_cot.add_heading(
                "2. Garantía de Continuidad (Reemplazo Inmediato)", level=3
            )
            doc_cot.add_paragraph(
                "Nos comprometemos a mantener la cobertura del servicio al 100%. En caso de ausencias por enfermedad, trámites administrativos o causas de fuerza mayor, el elemento será sustituido en un periodo no mayor a 90 minutos por personal de nuestro equipo de retén[cite: 10]."
            )

            doc_cot.add_heading("3. Confidencialidad Rigurosa", level=3)
            doc_cot.add_paragraph(
                "Todo el personal asignado cuenta con contratos de confidencialidad vigentes. Nuestra empresa se obliga a no divulgar información sensible, procesos internos o vulnerabilidades detectadas en las instalaciones del cliente[cite: 10]."
            )

            doc_cot.add_heading("4. Vigencia de la Oferta", level=3)
            doc_cot.add_paragraph(
                "La presente propuesta económica tiene una validez de 15 días naturales a partir de su fecha de emisión, debido a posibles ajustes en tabuladores de costos operativos[cite: 10]."
            )

            doc_cot.add_heading("5. Condiciones de Pago", level=3)
            doc_cot.add_paragraph(
                "Los servicios serán facturados de manera mensual y deberán ser liquidados dentro de los primeros 5 días naturales de cada mes para garantizar el flujo operativo y el cumplimiento puntual de sueldos del personal[cite: 10]."
            )

            doc_cot.add_paragraph(
                '"Nuestra estructura operativa garantiza que el error humano se reduzca al mínimo mediante la supervisión cruzada y el respaldo tecnológico en tiempo real."'
            )

            buffer_cot = BytesIO()
            doc_cot.save(buffer_cot)
            buffer_cot.seek(0)

            st.download_button(
                label="📥 Clic aquí para descargar la Cotización en Word",
                data=buffer_cot,
                file_name=f"Cotizacion_{empresa_cliente.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
