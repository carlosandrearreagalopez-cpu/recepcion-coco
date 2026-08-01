import streamlit as st
import pandas as pd
import os
from datetime import datetime
from streamlit_drawable_canvas import st_canvas
import io

# Configuración inicial
st.set_page_config(page_title="Recepción de Coco - LIF Brands", layout="wide")

# Estilos CSS de LIF Brands (Omitido parcialmente por brevedad, usa el mismo de tu guía)
st.markdown("<style>.stApp { background-color: #FFFFFF !important; }</style>", unsafe_allow_html=True)

FOTOS_DIR = "fotos_recepcion"
FIRMAS_DIR = "firmas_recepcion"
for d in [FOTOS_DIR, FIRMAS_DIR]:
    if not os.path.exists(d): os.makedirs(d)

EXCEL_FILE = "registros_recepcion_coco.xlsx"

def mostrar_logo(ancho=160):
    if os.path.exists("logo.png"):
        st.image("logo.png", width=ancho)

def cargar_datos():
    if os.path.exists(EXCEL_FILE):
        return pd.read_excel(EXCEL_FILE)
    return pd.DataFrame()

def guardar_datos(df):
    df.to_excel(EXCEL_FILE, index=False)

# Variables de estado
if "nav_state" not in st.session_state: st.session_state["nav_state"] = "home"

# --- NAVEGACIÓN ---
if st.session_state["nav_state"] == "home":
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        mostrar_logo(200)
        st.markdown("<h1 style='text-align: center; color: #1e3a8a;'>Control de Recepción de Coco</h1>", unsafe_allow_html=True)
        if st.button("Colaborador: Reportar nuevo ingreso", use_container_width=True, type="primary"):
            st.session_state["nav_state"] = "form"
            st.rerun()
        if st.button("Jefe de Calidad: Panel de Administración", use_container_width=True):
            st.session_state["nav_state"] = "admin"
            st.rerun()

# --- FORMULARIO DEL COLABORADOR ---
elif st.session_state["nav_state"] == "form":
    if st.button("Volver al inicio"):
        st.session_state["nav_state"] = "home"
        st.rerun()
    
    st.title("Nuevo Registro: Recepción de Coco")
    
    with st.form("form_coco"):
        # Encabezado basado en la imagen
        c1, c2, c3 = st.columns(3)
        with c1:
            responsable = st.text_input("Nombre del responsable")
            desc_materia = st.text_input("Descripción de materia prima")
        with c2:
            fecha = st.date_input("Fecha")
            proveedor = st.text_input("Nombre del proveedor")
            total_fruta = st.number_input("Total de Fruta Ingresada Planta", min_value=0.0)
        with c3:
            hora = st.time_input("Hora")
            observaciones = st.text_area("Observaciones")
            cant_unidades = st.number_input("Cantidad Unidades (Muestra)", min_value=0.0)
        
        st.subheader("Parámetros Fisicoquímicos")
        
        # Muestras
        muestras = {}
        for i in range(1, 4):
            st.markdown(f"**Muestra {i}**")
            mc1, mc2, mc3, mc4 = st.columns(4)
            with mc1: muestras[f"unidades_galon_{i}"] = st.number_input(f"Unidades/Galón (M{i})", key=f"ug_{i}")
            with mc2: muestras[f"volumen_{i}"] = st.number_input(f"Volumen Muestra (M{i})", key=f"v_{i}")
            with mc3: muestras[f"brix_{i}"] = st.number_input(f"Brix° (5 - 5.9) (M{i})", key=f"b_{i}")
            with mc4: muestras[f"ph_{i}"] = st.number_input(f"pH (M{i})", key=f"ph_{i}")

        submitted = st.form_submit_button("Enviar Registro a Revisión", type="primary")
        
        if submitted:
            nuevo_registro = {
                "ID_Registro": datetime.now().strftime("%Y%m%d%H%M%S"),
                "Estado": "Pendiente", # Estado inicial
                "Responsable": responsable, "Fecha": str(fecha), "Hora": str(hora),
                "Desc_Materia": desc_materia, "Observaciones": observaciones,
                "Proveedor": proveedor, "Total_Fruta": total_fruta, "Cant_Unidades": cant_unidades,
                **muestras,
                "Firma_Jefe": "Sin firma"
            }
            df = cargar_datos()
            df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
            guardar_datos(df)
            st.success("Registro enviado exitosamente. Pendiente de validación por el Jefe de Calidad.")

# --- PANEL DE ADMINISTRADOR (JEFE DE CALIDAD) ---
elif st.session_state["nav_state"] == "admin":
    if st.button("Volver al inicio"):
        st.session_state["nav_state"] = "home"
        st.rerun()
        
    st.title("Panel de Administrador - Jefe de Calidad")
    df = cargar_datos()
    
    if df.empty:
        st.info("No hay registros actualmente.")
    else:
        tab_pendientes, tab_aprobados, tab_todos = st.tabs(["Pendientes por Validar", "Aprobados", "Total de Registros"])
        
        # PESTAÑA 1: PENDIENTES
        with tab_pendientes:
            df_pendientes = df[df["Estado"] == "Pendiente"]
            if df_pendientes.empty:
                st.success("No hay registros pendientes de firma.")
            else:
                for idx, row in df_pendientes.iterrows():
                    with st.expander(f"Registro Pendiente: {row['Proveedor']} - {row['Fecha']}"):
                        st.write(row.drop(["Firma_Jefe", "Estado", "ID_Registro"]))
                        st.markdown("**Firma del Jefe de Calidad**")
                        canvas_result = st_canvas(
                            fill_color="rgba(255, 255, 255, 0.3)", stroke_width=2,
                            stroke_color="#000000", background_color="#EEEEEE",
                            height=100, width=400, drawing_mode="freedraw",
                            key=f"firma_{row['ID_Registro']}"
                        )
                        if st.button("Validar y Firmar", key=f"btn_{row['ID_Registro']}"):
                            if canvas_result.image_data is not None:
                                import numpy as np
                                from PIL import Image
                                img = Image.fromarray(canvas_result.image_data.astype('uint8'), mode="RGBA")
                                nombre_firma = f"firma_{row['ID_Registro']}.png"
                                img.save(os.path.join(FIRMAS_DIR, nombre_firma))
                                
                                # Actualizar estado en excel
                                df.at[idx, "Estado"] = "Aprobado"
                                df.at[idx, "Firma_Jefe"] = nombre_firma
                                guardar_datos(df)
                                st.success("Registro validado exitosamente.")
                                st.rerun()

        # PESTAÑA 2: APROBADOS
        with tab_aprobados:
            df_aprobados = df[df["Estado"] == "Aprobado"]
            st.dataframe(df_aprobados)
            
        # PESTAÑA 3: TODOS LOS REGISTROS
        with tab_todos:
            filtro_fecha = st.date_input("Filtrar por fecha", value=None)
            df_filtrado = df.copy()
            if filtro_fecha:
                df_filtrado = df_filtrado[df_filtrado["Fecha"] == str(filtro_fecha)]
                
            st.dataframe(df_filtrado[["ID_Registro", "Fecha", "Proveedor", "Estado"]])
            
            st.markdown("### Descargar Registros Llenos")
            for idx, row in df_filtrado.iterrows():
                if row["Estado"] == "Aprobado":
                    # Aquí irá la lógica de integración con el PDF
                    if st.button(f"Descargar PDF Lleno - {row['ID_Registro']}", key=f"dl_{row['ID_Registro']}"):
                        st.info("El documento se generará usando el formato PDF que compartirás.")
                        # generar_pdf_relleno(row) -> Función a implementar con tu PDF
