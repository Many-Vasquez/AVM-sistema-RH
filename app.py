import streamlit as st
import pandas as pd
from datetime import datetime

# Configuración de la página
st.set_page_config(page_title="AVM Seguridad - Gestión de Personal", page_icon="🛡️", layout="wide")

# Título principal
st.title("🛡️ AVM GRUPO INTEGRAL DE SEGURIDAD PRIVADA DEL NORTE")
st.subheader("Sistema Automatizado de Recursos Humanos y Contratos")

# Menú lateral de navegación
menu = st.sidebar.selectbox("Menú Principal", ["📊 Dashboard / Empleados", "➕ Nuevo Registro & Contrato"])

# Base de datos simulada en memoria (puedes conectarla a Google Sheets después)
if 'empleados' not in st.session_state:
    st.session_state['empleados'] = pd.DataFrame(columns=[
        "ID", "Nombre Completo", "CURP", "Puesto", "Salario Diario", "Fecha de Alta"
    ])

if menu == "📊 Dashboard / Empleados":
    st.markdown("### Listado de Personal Activo")
    
    if st.session_state['empleados'].empty:
        st.info("No hay empleados registrados todavía. Utiliza la sección 'Nuevo Registro & Contrato' para agregar al primero.")
    else:
        st.dataframe(st.session_state['empleados'], use_container_width=True)
        
        # Métricas rápidas
        total_empleados = len(st.session_state['empleados'])
        st.metric(label="Total de Elementos Registrados", value=total_empleados)

elif menu == "➕ Nuevo Registro & Contrato":
    st.markdown("### Alta de Nuevo Elemento y Cálculo IMSS / Contrato")
    
    with st.form("form_empleado"):
        col1, col2 = st.columns(2)
        
        with col1:
            nombre = st.text_input("Nombre Completo del Trabajador")
            curp = st.text_input("CURP")
            puesto = st.selectbox("Puesto", ["Guardia Intramuros", "Supervisor de Operaciones", "Custodio"])
            
        with col2:
            salario_diario = st.number_input("Salario Diario (MXN)", min_value=250.0, value=300.0, step=10.0)
            fecha_alta = st.date_input("Fecha de Ingreso", datetime.now())
            
        submitted = st.form_submit_button("Guardar Empleado y Calcular Prestaciones")
        
        if submitted:
            if nombre and curp:
                # Simulación del cálculo de integración de SBC (ejemplo base ley)
                factor_integracion = 1.0493 # Factor mínimo primer año (aguinaldo y prima vacacional)
                sbc = salario_diario * factor_integracion
                
                nuevo_registro = pd.DataFrame({
                    "ID": [len(st.session_state['empleados']) + 1],
                    "Nombre Completo": [nombre],
                    "CURP": [curp],
                    "Puesto": [puesto],
                    "Salario Diario": [f"${salario_diario:,.2f}"],
                    "Fecha de Alta": [str(fecha_alta)]
                })
                
                st.session_state['empleados'] = pd.concat([st.session_state['empleados'], nuevo_registro], ignore_index=True)
                
                st.success(f"¡Empleado {nombre} registrado exitosamente!")
                st.info(f"Salario Base de Cotización (SBC) calculado para el IMSS: **${sbc:,.2f}**")
            else:
                st.error("Por favor completa al menos el nombre y la CURP.")
