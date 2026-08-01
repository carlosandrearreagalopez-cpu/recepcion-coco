import streamlit as st
import pandas as pd
import os
from datetime import datetime
from streamlit_drawable_canvas import st_canvas
import fitz  # PyMuPDF para manejar el PDF

# Configuración inicial de la página
st.set_page_config(page_title="Control de Recepción de Coco - LIF Brands", layout="wide")

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
.card-registro {
    padding: 15px;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    margin-bottom: 12px;
    background-color: #f8fafc;
    border-left: 5px solid #1e3a8a;
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

# Función optimizada para rellenar el PDF en las coordenadas exactas de las celdas
def generar_pdf_relleno(registro):
    if not os.path.exists(PDF_PLANTILLA):
        return None
    
    doc = fitz.open(PDF_PLANTILLA)
    pagina = doc[0]
    
    # Coordenadas exactas alineadas con el formato oficial de la imagen
    coordenadas = {
        "Responsable": (200, 150),
        "Fecha": (460, 150),
        "Hora": (650, 150),
        "Desc_Materia": (200, 178),
        "Observaciones": (500, 185),
        "Proveedor": (200, 205),
        "Total_Fruta": (200, 235),
        "Cant_Unidades": (200, 260),
        
        # Muestra 1
        "unidades_galon_1": (200, 325),
        "volumen_1": (200, 350),
        "brix_1": (200, 375),
        "ph_1": (200, 400),
        
        # Muestra 2
        "unidades_galon_2": (200, 440),
        "volumen_2": (200, 465),
        "brix_2": (200, 490),
        "ph_2": (200, 515),
        
        # Muestra 3
        "unidades_galon_3": (200, 555),
        "volumen_3": (200, 580),
        "brix_3": (200, 605),
        "ph_3": (200, 630),
    }

    for campo, coord in coordenadas.items():
        valor = str(registro.get(campo, ""))
        pagina.insert_text(coord, valor, fontsize=9, color=(0, 0, 0))

    nombre_firma = registro.get("Firma_Jefe", "Sin firma")
    ruta_firma = os.path.join(FIRMAS_DIR, str(nombre_firma))
    
    if nombre_firma != "Sin firma" and os.path.exists(ruta_firma):
        rectangulo_firma = fitz.Rect(230, 675, 380, 725)
        pagina.insert_image(rectangulo_firma, filename=ruta_firma)

    pdf_bytes = doc.write()
    doc.close()
    return pdf_bytes

# Control de navegación y estados de sesión
if "nav_state" not in st.session_state: 
    st.session_state["nav_state"] = "home"
if "form_logueado" not in st.session_state:
    st.session_state["form_logueado"] = False
if "admin_logueado" not in st.session_state:
    st.session_state["admin_logueado"] = False
if "enviado_exitoso" not in st.session_state:
    st.session_state["enviado_exitoso"] = False

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
            st.session_state["nav_state"] = "form_login"
            st.rerun()
            
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔒 Jefe de Calidad: Panel de Administración", use_container_width=True):
            st.session_state["nav_state"] = "admin_login"
            st.rerun()

# ==========================================
# 2. LOGIN PARA EL NUEVO INGRESO (COLABORADOR)
# ==========================================
elif st.session_state["nav_state"] == "form_login":
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("⬅️ Volver al inicio"):
            st.session_state["nav_state"] = "home"
            st.rerun()
        st.title("Acceso a Registro")
        st.markdown("Ingrese la contraseña autorizada (`20lf26`) para reportar un nuevo ingreso:")
        
        password_form = st.text_input("Contraseña de ingreso", type="password")
        if st.button("Verificar Acceso", use_container_width=True, type="primary"):
            if password_form == "20lf26":
                st.session_state["form_logueado"] = True
                st.session_state["nav_state"] = "form"
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")

# ==========================================
# 3. FORMULARIO DEL COLABORADOR
# ==========================================
elif st.session_state["nav_state"] == "form":
    if not st.session_state["form_logueado"]:
        st.session_state["nav_state"] = "form_login"
        st.rerun()
        
    if st.button("⬅️ Volver al inicio"):
        st.session_state["form_logueado"] = False
        st.session_state["nav_state"] = "home"
        st.rerun()
    
    mostrar_logo(140)
    st.title("Registro de Recepción de Coco")
    
    if st.session_state["enviado_exitoso"]:
        st.success("¡Registro enviado con éxito! Quedará pendiente de validación por el Jefe de Calidad.")
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            if st.button("➕ Ingresar un nuevo registro", use_container_width=True, type="primary"):
                st.session_state["enviado_exitoso"] = False
                st.rerun()
        with col_b2:
            if st.button("🏠 Volver al inicio", use_container_width=True):
                st.session_state["enviado_exitoso"] = False
                st.session_state["form_logueado"] = False
                st.session_state["nav_state"] = "home"
                st.rerun()
    else:
        # Proveedor fuera del form para reacción inmediata en celulares
        st.markdown("### Selección de Proveedor")
        proveedor_opcion = st.selectbox("Nombre del proveedor", ["GRANOS BASICOS LA PATRONA SOCIEDAD ANONIMA", "Otro"])
        proveedor_final = ""
        if proveedor_opcion == "Otro":
            proveedor_final = st.text_input("Escriba el nombre del nuevo proveedor:")

        st.markdown("---")

        with st.form("form_coco"):
            st.header("1. Datos Generales")
            c1, c2, c3 = st.columns(3)
            with c1:
                responsable = st.selectbox(
                    "Nombre del responsable", 
                    ["Carlos Canto", "Carlos Rodas", "Jonathan", "Damarias Arellanos", "Carlos López", "Marlon Escobar"]
                )
                desc_materia = st.text_input("Descripción de materia prima", value="Coco")
            with c2:
                fecha = st.date_input("Fecha")
                total_fruta = st.number_input("Total de Fruta Ingresada Planta", min_value=0.0, value=0.0)
            with c3:
                hora = st.time_input("Hora")
                observaciones = st.text_area("Observaciones", value="Ninguna")
                cant_unidades = st.number_input("Cantidad Unidades (Muestra)", min_value=0.0, value=0.0)
            
            st.header("2. Parámetros Fisicoquímicos")
            
            st.subheader("Muestra 1")
            mc1_1, mc1_2, mc1_3, mc1_4 = st.columns(4)
            with mc1_1: ug_1 = st.number_input("Unidades/Galón (M1)", min_value=0.0, value=0.0, key="ug_1")
            with mc1_2: v_1 = st.number_input("Volumen (M1)", min_value=0.0, value=0.0, key="v_1")
            with mc1_3: b_1 = st.number_input("Brix° (5-5.9) (M1)", min_value=0.0, value=0.0, key="b_1", format="%.2f")
            with mc1_4: ph_1 = st.number_input("pH (M1)", min_value=0.0, value=0.0, key="ph_1", format="%.2f")

            st.subheader("Muestra 2")
            mc2_1, mc2_2, mc2_3, mc2_4 = st.columns(4)
            with mc2_1: ug_2 = st.number_input("Unidades/Galón (M2)", min_value=0.0, value=0.0, key="ug_2")
            with mc2_2: v_2 = st.number_input("Volumen (M2)", min_value=0.0, value=0.0, key="v_2")
            with mc2_3: b_2 = st.number_input("Brix° (5-5.9) (M2)", min_value=0.0, value=0.0, key="b_2", format="%.2f")
            with mc2_4: ph_2 = st.number_input("pH (M2)", min_value=0.0, value=0.0, key="ph_2", format="%.2f")

            st.subheader("Muestra 3")
            mc3_1, mc3_2, mc3_3, mc3_4 = st.columns(4)
            with mc3_1: ug_3 = st.number_input("Unidades/Galón (M3)", min_value=0.0, value=0.0, key="ug_3")
            with mc3_2: v_3 = st.number_input("Volumen (M3)", min_value=0.0, value=0.0, key="v_3")
            with mc3_3: b_3 = st.number_input("Brix° (5-5.9) (M3)", min_value=0.0, value=0.0, key="b_3", format="%.2f")
            with mc3_4: ph_3 = st.number_input("pH (M3)", min_value=0.0, value=0.0, key="ph_3", format="%.2f")

            submitted = st.form_submit_button("Guardar y Enviar a Revisión", type="primary")
            
            if submitted:
                proveedor_a_guardar = proveedor_final if proveedor_opcion == "Otro" else proveedor_opcion
                
                if proveedor_opcion == "Otro" and not proveedor_final.strip():
                    st.error("Por favor, ingrese el nombre del nuevo proveedor.")
                else:
                    nuevo_registro = {
                        "ID_Registro": datetime.now().strftime("%Y%m%d%H%M%S"),
                        "Estado": "Pendiente",
                        "Responsable": responsable, 
                        "Fecha": str(fecha), 
                        "Hora": str(hora),
                        "Desc_Materia": desc_materia, 
                        "Observaciones": observaciones,
                        "Proveedor": proveedor_a_guardar, 
                        "Total_Fruta": total_fruta, 
                        "Cant_Unidades": cant_unidades,
                        "unidades_galon_1": ug_1, "volumen_1": v_1, "brix_1": b_1, "ph_1": ph_1,
                        "unidades_galon_2": ug_2, "volumen_2": v_2, "brix_2": b_2, "ph_2": ph_2,
                        "unidades_galon_3": ug_3, "volumen_3": v_3, "brix_3": b_3, "ph_3": ph_3,
                        "Firma_Jefe": "Sin firma"
                    }
                    df = cargar_datos()
                    df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
                    guardar_datos(df)
                    st.session_state["enviado_exitoso"] = True
                    st.rerun()

# ==========================================
# 4. LOGIN DE ADMINISTRADOR
# ==========================================
elif st.session_state["nav_state"] == "admin_login":
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("⬅️ Volver al inicio"):
            st.session_state["nav_state"] = "home"
            st.rerun()
            
        st.title("Panel de Administrador")
        st.markdown("Ingrese la contraseña de administrador (`glad726lif`) para acceder a los registros.")
        
        password_input = st.text_input("Contraseña de administrador", type="password")
        if st.button("Verificar Acceso", use_container_width=True, type="primary"):
            if password_input == "glad726lif":  
                st.session_state["admin_logueado"] = True
                st.session_state["nav_state"] = "admin_dashboard"
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")

# ==========================================
# 5. DASHBOARD DEL ADMINISTRADOR (ESTILO INSPECCIÓN)
# ==========================================
elif st.session_state["nav_state"] == "admin_dashboard":
    if not st.session_state.get("admin_logueado", False):
        st.session_state["nav_state"] = "admin_login"
        st.rerun()
        
    col_h1, col_h2 = st.columns([4, 1])
    with col_h1:
        mostrar_logo(120)
        st.header("Panel de Administrador - Jefe de Calidad")
    with col_h2:
        if st.button("Cerrar Sesión"):
            st.session_state["admin_logueado"] = False
            st.session_state["nav_state"] = "home"
            st.rerun()
            
    df = cargar_datos()
    if df.empty:
        st.info("No hay registros guardados en el sistema actualmente.")
    else:
        # Pestañas limpias estilo sistema corporativo
        tab_pendientes, tab_aprobados, tab_todos = st.tabs([
            "⏳ Pendientes por Validar", 
            "✅ Aprobados", 
            "📊 Historial Completo de Registros"
        ])
        
        with tab_pendientes:
            df_pendientes = df[df["Estado"] == "Pendiente"]
            if df_pendientes.empty:
                st.success("¡Excelente! No hay registros pendientes de firma.")
            else:
                for idx, row in df_pendientes.iterrows():
                    with st.container():
                        st.markdown(f"""
                        <div class="card-registro">
                        <b>#{row['ID_Registro']} — Proveedor: {row['Proveedor']}</b><br>
                        Fecha: {row['Fecha']} | Responsable: {row['Responsable']} | Total Fruta: {row['Total_Fruta']}
                        </div>
                        """, unsafe_allow_html=True)
                        
                        if st.button(f"Revisar y Firmar #{row['ID_Registro']}", key=f"btn_ver_{idx}"):
                            st.write(row.to_dict())
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
                            
                            if st.button(f"Validar y Firmar Definitivamente #{row['ID_Registro']}", key=f"btn_firmar_{row['ID_Registro']}", type="primary"):
                                if canvas_result.image_data is not None:
                                    import numpy as np
                                    from PIL import Image
                                    img_data = canvas_result.image_data
                                    img = Image.fromarray(img_data.astype('uint8'), mode="RGBA")
                                    nombre_firma = f"firma_{row['ID_Registro']}.png"
                                    ruta_firma = os.path.join(FIRMAS_DIR, nombre_firma)
                                    img.save(ruta_firma)
                                    
                                    df.at[idx, "Estado"] = "Aprobado"
                                    df.at[idx, "Firma_Jefe"] = nombre_firma
                                    guardar_datos(df)
                                    st.success("¡Registro validado y firmado correctamente!")
                                    st.rerun()
                                else:
                                    st.warning("Por favor dibuje la firma antes de validar.")

        with tab_aprobados:
            df_aprobados = df[df["Estado"] == "Aprobado"]
            if df_aprobados.empty:
                st.info("Aún no hay registros aprobados.")
            else:
                for idx, row in df_aprobados.iterrows():
                    st.markdown(f"""
                    <div class="card-registro">
                    <b>✅ ID: {row['ID_Registro']} | Proveedor: {row['Proveedor']}</b><br>
                    Fecha: {row['Fecha']} | Responsable: {row['Responsable']} | Estado: Aprobado
                    </div>
                    """, unsafe_allow_html=True)

        with tab_todos:
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
                estado_badge = "🟢 Aprobado" if row["Estado"] == "Aprobado" else "🟠 Pendiente"
                with st.container():
                    st.markdown(f"""
                    <div class="card-registro">
                    <b>ID: {row['ID_Registro']}</b> | <b>Fecha:</b> {row['Fecha']} | <b>Proveedor:</b> {row['Proveedor']} | <b>Estado:</b> {estado_badge}
                    </div>
                    """, unsafe_allow_html=True)
                    
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
