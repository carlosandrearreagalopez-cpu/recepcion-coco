import streamlit as st
import pandas as pd
import os
from datetime import datetime
from streamlit_drawable_canvas import st_canvas
import fitz  # PyMuPDF para manejar el PDF

# Configuración inicial de la página
st.set_page_config(page_title="Recepción de Coco - LIF Brands", layout="wide")

# ==========================================
# ESTILOS CSS CON IDENTIDAD VISUAL LIF BRANDS
# ==========================================
st.markdown("""
<style>
.stApp { background-color: #FFFFFF !important; }
html, body, [class*="css"], p, span, label { font-family: Arial, sans-serif !important; color: #000000 !important; }
h1, h2, h3, h4, h5, h6 { color: #1e3a8a !important; font-family: Arial, sans-serif !important; }
.stTextInput label, .stSelectbox label, .stDateInput label, .stNumberInput label, .stRadio label, .stFileUploader label {
    color: #1e3a8a !important;
    font-weight: bold !important;
}
.stButton>button {
    background-color: #FFFFFF !important;
    color: #1e3a8a !important;
    border: 2px solid #1e3a8a !important;
    border-radius: 6px;
    font-family: Arial, sans-serif;
    font-weight: bold;
}
.stButton>button:hover {
    background-color: #1e3a8a !important;
    color: #FFFFFF !important;
}
button[kind="primary"] {
    background-color: #1e3a8a !important;
    color: #FFFFFF !important;
    border: none !important;
}
button[kind="primary"]:hover {
    background-color: #3b82f6 !important;
}
input, select {
    background-color: #f8fafc !important;
    color: #000000 !important;
    border: 1px solid #cbd5e1 !important;
    border-radius: 4px;
}
</style>
""", unsafe_allow_html=True)

# Directorios necesarios
FIRMAS_DIR = "firmas_recepcion"
if not os.path.exists(FIRMAS_DIR):
    os.makedirs(FIRMAS_DIR)

EXCEL_FILE = "registros_recepcion_coco.xlsx"
PDF_PLANTILLA = "R ICC15.2 - Control de recepción Coco.pdf"

def mostrar_logo(ancho=160):
    if os.path.exists("logo.png"):
        st.image("logo.png", width=ancho)
    else:
        st.warning("⚠️ Logo no encontrado. Asegúrate de tener 'logo.png' en el repositorio.")

def cargar_datos():
    if os.path.exists(EXCEL_FILE):
        return pd.read_excel(EXCEL_FILE)
    return pd.DataFrame()

def guardar_datos(df):
    df.to_excel(EXCEL_FILE, index=False)

# Función para rellenar el PDF y colocar la firma digital encima del texto "Jefe de Calidad"
def generar_pdf_relleno(registro):
    if not os.path.exists(PDF_PLANTILLA):
        return None
    
    doc = fitz.open(PDF_PLANTILLA)
    pagina = doc[0]
    
    # Coordenadas (X, Y) basadas en el formato PDF oficial
    coordenadas = {
        "Responsable": (180, 150),
        "Fecha": (450, 150),
        "Hora": (650, 150),
        "Desc_Materia": (180, 175),
        "Observaciones": (500, 175),
        "Proveedor": (180, 200),
        "Total_Fruta": (180, 225),
        "Cant_Unidades": (180, 250),
        
        # Muestra 1
        "unidades_galon_1": (180, 310),
        "volumen_1": (180, 330),
        "brix_1": (180, 350),
        "ph_1": (180, 370),
        
        # Muestra 2
        "unidades_galon_2": (180, 420),
        "volumen_2": (180, 440),
        "brix_2": (180, 460),
        "ph_2": (180, 480),
        
        # Muestra 3
        "unidades_galon_3": (180, 530),
        "volumen_3": (180, 550),
        "brix_3": (180, 570),
        "ph_3": (180, 590),
    }

    # Escribir textos en el PDF
    for campo, coord in coordenadas.items():
        valor = str(registro.get(campo, ""))
        pagina.insert_text(coord, valor, fontsize=9, color=(0, 0, 0))

    # Insertar la firma del Jefe de Calidad si existe
    nombre_firma = registro.get("Firma_Jefe", "Sin firma")
    ruta_firma = os.path.join(FIRMAS_DIR, str(nombre_firma))
    
    if nombre_firma != "Sin firma" and os.path.exists(ruta_firma):
        # Rectángulo ubicado justo encima de la línea "Jefe de Calidad" al pie de página
        rectangulo_firma = fitz.Rect(230, 680, 380, 730)
        pagina.insert_image(rectangulo_firma, filename=ruta_firma)

    pdf_bytes = doc.write()
    doc.close()
    return pdf_bytes

# Control de navegación en la sesión
if "nav_state" not in st.session_state: 
    st.session_state["nav_state"] = "home"

# ==========================================
# 1. PANTALLA DE INICIO
# ==========================================
if st.session_state["nav_state"] == "home":
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        mostrar_logo(200)
        st.markdown("<h1 style='text-align: center; color: #1e3a8a;'>Control de Recepción de Coco</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #65a30d; font-weight: bold;'>LIF Brands Aseguramiento de Calidad</p>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("📝 Colaborador: Reportar nuevo ingreso", use_container_width=True, type="primary"):
            st.session_state["nav_state"] = "form"
            st.rerun()
            
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔒 Jefe de Calidad: Panel de Administración", use_container_width=True):
            st.session_state["nav_state"] = "admin_login"
            st.rerun()

# ==========================================
# 2. FORMULARIO DEL COLABORADOR
# ==========================================
elif st.session_state["nav_state"] == "form":
    if st.button("⬅️ Volver al inicio"):
        st.session_state["nav_state"] = "home"
        st.rerun()
    
    mostrar_logo(140)
    st.title("Registro de Recepción de Coco")
    
    with st.form("form_coco"):
        st.header("1. Datos Generales")
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
        
        st.header("2. Parámetros Fisicoquímicos")
        muestras = {}
        for i in range(1, 4):
            st.subheader(f"Muestra {i}")
            mc1, mc2, mc3, mc4 = st.columns(4)
            with mc1: muestras[f"unidades_galon_{i}"] = st.number_input(f"Unidades/Galón (M{i})", key=f"ug_{i}")
            with mc2: muestras[f"volumen_{i}"] = st.number_input(f"Volumen (M{i})", key=f"v_{i}")
            with mc3: muestras[f"brix_{i}"] = st.number_input(f"Brix° (5-5.9) (M{i})", key=f"b_{i}", format="%.2f")
            with mc4: muestras[f"ph_{i}"] = st.number_input(f"pH (M{i})", key=f"ph_{i}", format="%.2f")

        submitted = st.form_submit_button("Guardar y Enviar a Revisión", type="primary")
        
        if submitted:
            if not responsable.strip() or not proveedor.strip():
                st.error("Por favor complete el nombre del responsable y del proveedor.")
            else:
                nuevo_registro = {
                    "ID_Registro": datetime.now().strftime("%Y%m%d%H%M%S"),
                    "Estado": "Pendiente",
                    "Responsable": responsable, 
                    "Fecha": str(fecha), 
                    "Hora": str(hora),
                    "Desc_Materia": desc_materia, 
                    "Observaciones": observaciones,
                    "Proveedor": proveedor, 
                    "Total_Fruta": total_fruta, 
                    "Cant_Unidades": cant_unidades,
                    **muestras,
                    "Firma_Jefe": "Sin firma"
                }
                df = cargar_datos()
                df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
                guardar_datos(df)
                st.success("¡Registro enviado con éxito! Quedará pendiente de validación por el Jefe de Calidad.")

# ==========================================
# 3. LOGIN DE ADMINISTRADOR
# ==========================================
elif st.session_state["nav_state"] == "admin_login":
    if st.button("⬅️ Volver al inicio"):
        st.session_state["nav_state"] = "home"
        st.rerun()
        
    st.title("Acceso - Panel de Administrador")
    password_input = st.text_input("Contraseña de Administrador", type="password")
    
    if st.button("Verificar Acceso", type="primary"):
        if password_input == "glad726lif":  # Contraseña heredada de tu estándar
            st.session_state["admin_logueado"] = True
            st.session_state["nav_state"] = "admin_dashboard"
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")

# ==========================================
# 4. DASHBOARD DEL ADMINISTRADOR
# ==========================================
elif st.session_state["nav_state"] == "admin_dashboard":
    if not st.session_state.get("admin_logueado", False):
        st.session_state["nav_state"] = "admin_login"
        st.rerun()
        
    if st.button("⬅️ Cerrar Sesión / Volver al inicio"):
        st.session_state["admin_logueado"] = False
        st.session_state["nav_state"] = "home"
        st.rerun()
        
    mostrar_logo(140)
    st.header("Panel de Administrador - Jefe de Calidad")
    
    df = cargar_datos()
    if df.empty:
        st.info("No hay registros guardados en el sistema actualmente.")
    else:
        # Pestañas requeridas por ti
        tab_pendientes, tab_aprobados, tab_todos = st.tabs([
            "📌 Registros Pendientes por Validar", 
            "✅ Registros Aprobados", 
            "📊 Total de Registros (Filtros)"
        ])
        
        # PESTAÑA 1: PENDIENTES DE VALIDAR Y FIRMAR
        with tab_pendientes:
            df_pendientes = df[df["Estado"] == "Pendiente"]
            if df_pendientes.empty:
                st.success("¡Excelente! No hay registros pendientes de firma.")
            else:
                for idx, row in df_pendientes.iterrows():
                    with st.expander(f"Ingreso #{row['ID_Registro']} | Proveedor: {row['Proveedor']} | Fecha: {row['Fecha']}"):
                        st.write(row.to_dict())
                        st.markdown("---")
                        st.markdown("#### Dibuje su firma de aprobación:")
                        
                        canvas_result = st_canvas(
                            fill_color="rgba(255, 255, 255, 0.3)", 
                            stroke_width=2,
                            stroke_color="#1e3a8a", 
                            background_color="#FFFFFF",
                            height=120, 
                            width=400, 
                            drawing_mode="freedraw",
                            key=f"canvas_firma_{row['ID_Registro']}"
                        )
                        
                        if st.button(f"Validar y Firmar Registro #{row['ID_Registro']}", key=f"btn_firmar_{row['ID_Registro']}", type="primary"):
                            if canvas_result.image_data is not None:
                                import numpy as np
                                from PIL import Image
                                img_data = canvas_result.image_data
                                img = Image.fromarray(img_data.astype('uint8'), mode="RGBA")
                                nombre_firma = f"firma_{row['ID_Registro']}.png"
                                ruta_firma = os.path.join(FIRMAS_DIR, nombre_firma)
                                img.save(ruta_firma)
                                
                                # Actualizar DataFrame
                                df.at[idx, "Estado"] = "Aprobado"
                                df.at[idx, "Firma_Jefe"] = nombre_firma
                                guardar_datos(df)
                                st.success("¡Registro validado y firmado correctamente!")
                                st.rerun()
                            else:
                                st.warning("Por favor dibuje la firma antes de validar.")

        # PESTAÑA 2: APROBADOS
        with tab_aprobados:
            st.subheader("Registros aprobados y listos")
            df_aprobados = df[df["Estado"] == "Aprobado"]
            if df_aprobados.empty:
                st.info("Aún no hay registros aprobados.")
            else:
                st.dataframe(df_aprobados[["ID_Registro", "Fecha", "Proveedor", "Responsable", "Estado"]], use_container_width=True)

        # PESTAÑA 3: TOTAL DE REGISTROS (CON FILTROS Y DESCARGA)
        with tab_todos:
            st.subheader("Historial Completo de Registros")
            
            # Filtros por fecha y proveedor
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                filtro_fecha = st.date_input("Filtrar por Fecha (Opcional)", value=None)
            with col_f2:
                lista_prov = ["Todos"] + list(df["Proveedor"].dropna().unique())
                filtro_proveedor = st.selectbox("Filtrar por Proveedor", lista_prov)
                
            df_filtrado = df.copy()
            if filtro_fecha:
                df_filtrado = df_filtrado[df_filtrado["Fecha"] == str(filtro_fecha)]
            if filtro_proveedor != "Todos":
                df_filtrado = df_filtrado[df_filtrado["Proveedor"] == filtro_proveedor]
                
            st.markdown(f"Mostrando **{len(df_filtrado)}** registros:")
            
            for idx, row in df_filtrado.iterrows():
                estado_color = "🟢 Aprobado" if row["Estado"] == "Aprobado" else "🟠 Pendiente"
                with st.container():
                    st.markdown(f"""
                    **ID:** {row['ID_Registro']} | **Fecha:** {row['Fecha']} | **Proveedor:** {row['Proveedor']} | **Estado:** {estado_color}
                    """, unsafe_allow_html=True)
                    
                    # Si ya está aprobado, permitir descargar el PDF relleno con la firma oficial
                    if row["Estado"] == "Aprobado":
                        pdf_bytes = generar_pdf_relleno(row.to_dict())
                        if pdf_bytes:
                            st.download_button(
                                label=f"📥 Descargar PDF Relleno y Firmado (#{row['ID_Registro']})",
                                data=pdf_bytes,
                                file_name=f"Control_Coco_{row['ID_Registro']}.pdf",
                                mime="application/pdf",
                                key=f"download_{row['ID_Registro']}"
                            )
                    st.markdown("---")
