import streamlit as st
import pandas as pd
import os
from datetime import datetime
from streamlit_drawable_canvas import st_canvas
import io
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas as pdf_canvas

# Configuración inicial de la página
st.set_page_config(page_title="Control de Recepción - LIF Brands", layout="wide")

# ==========================================
# ESTILOS CSS CORREGIDOS (Garantiza lectura de texto)
# ==========================================
st.markdown("""
<style>
.stApp { background-color: #FFFFFF !important; }
html, body, [class*="css"], p, span, label { font-family: Arial, sans-serif !important; color: #1e293b !important; }
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
    color: #0f172a !important;
    border: 1px solid #cbd5e1 !important;
    border-radius: 4px;
}
/* Tarjetas del panel administrador con texto oscuro forzado */
.card-seccion {
    padding: 15px;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    margin-bottom: 15px;
    background-color: #f1f5f9 !important;
    border-left: 5px solid #1e3a8a;
    color: #0f172a !important; 
}
.card-seccion b, .card-seccion span, .card-seccion p {
    color: #0f172a !important;
}
/* Estilo para expanders */
.streamlit-expanderHeader {
    color: #1e3a8a !important;
    font-weight: bold !important;
}
</style>
""", unsafe_allow_html=True)

# Directorios necesarios
FIRMAS_DIR = "firmas_recepcion"
if not os.path.exists(FIRMAS_DIR):
    os.makedirs(FIRMAS_DIR)

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

def eliminar_registro(id_registro):
    df = cargar_datos()
    df = df[df["ID_Registro"] != id_registro]
    guardar_datos(df)
    st.success(f"Registro #{id_registro} eliminado correctamente.")

# ==========================================
# GENERADOR DE PDF DESDE CERO (Formato Limpio)
# ==========================================
def generar_pdf_nuevo(registro):
    buffer = io.BytesIO()
    c = pdf_canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    
    # 1. Encabezado
    if os.path.exists("logo.png"):
        c.drawImage("logo.png", 40, height - 80, width=120, height=50, preserveAspectRatio=True, mask='auto')
    
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width / 2.0, height - 50, "Control de Recepción de Coco")
    
    c.setFont("Helvetica", 9)
    c.drawString(width - 160, height - 40, "Código: R ICC/15-2")
    c.drawString(width - 160, height - 52, "Versión: 02")
    c.drawString(width - 160, height - 64, f"ID de Sistema: #{registro.get('ID_Registro', '')}")
    
    c.line(40, height - 90, width - 40, height - 90)
    
    # 2. Datos Generales (Estructura tipo tabla)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, height - 110, "Datos Generales")
    
    c.setFont("Helvetica", 10)
    y = height - 130
    c.drawString(40, y, f"Responsable: {registro.get('Responsable', '')}")
    c.drawString(300, y, f"Fecha: {registro.get('Fecha', '')}")
    c.drawString(460, y, f"Hora: {registro.get('Hora', '')}")
    
    y -= 20
    c.drawString(40, y, f"Materia Prima: {registro.get('Desc_Materia', '')}")
    c.drawString(300, y, f"Total de Fruta Ingresada: {registro.get('Total_Fruta', '')}")
    
    y -= 20
    c.drawString(40, y, f"Proveedor: {registro.get('Proveedor', '')}")
    c.drawString(300, y, f"Cantidad Unidades (Muestra): {registro.get('Cant_Unidades', '')}")
    
    y -= 20
    c.drawString(40, y, f"Observaciones: {registro.get('Observaciones', 'Ninguna')}")
    
    c.line(40, y - 10, width - 40, y - 10)
    
    # 3. Parámetros Fisicoquímicos (Tabla)
    y -= 30
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "Parámetros Fisicoquímicos")
    
    y -= 25
    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, y, "Métricas")
    c.drawString(200, y, "Muestra 1")
    c.drawString(320, y, "Muestra 2")
    c.drawString(440, y, "Muestra 3")
    
    y -= 10
    c.line(40, y, width - 40, y)
    
    y -= 15
    c.setFont("Helvetica", 10)
    parametros = ["Unidades/Galón", "Volumen", "Brix° (5 - 5.9)", "pH"]
    keys = ["unidades_galon", "volumen", "brix", "ph"]
    
    for param, key in zip(parametros, keys):
        c.drawString(40, y, param)
        c.drawString(200, y, str(registro.get(f"{key}_1", "")))
        c.drawString(320, y, str(registro.get(f"{key}_2", "")))
        c.drawString(440, y, str(registro.get(f"{key}_3", "")))
        y -= 20
    
    c.line(40, y + 5, width - 40, y + 5)
    
    # 4. Firmas y Estado
    y -= 30
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "Aprobación y Resolución")
    
    y -= 20
    c.setFont("Helvetica", 10)
    c.drawString(40, y, f"Estado del Registro: {registro.get('Estado', '')}")
    
    nombre_firma = registro.get("Firma_Jefe", "Sin firma")
    ruta_firma = os.path.join(FIRMAS_DIR, str(nombre_firma))
    
    if nombre_firma != "Sin firma" and os.path.exists(ruta_firma):
        c.drawImage(ruta_firma, 40, y - 60, width=120, height=45, preserveAspectRatio=True, mask='auto')
        c.line(40, y - 65, 180, y - 65)
        c.drawString(40, y - 80, "Firma de V°B° Calidad")
    
    c.save()
    buffer.seek(0)
    return buffer.getvalue()

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
        st.markdown("<h1 style='text-align: center;'>Control de Recepción de Coco</h1>", unsafe_allow_html=True)
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
        st.markdown("Ingrese la contraseña autorizada (`20lf26`):")
        
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
        st.success("¡Registro enviado con éxito! Quedará pendiente de validación.")
        c_btn1, c_btn2 = st.columns(2)
        with c_btn1:
            if st.button("➕ Ingresar un nuevo registro", use_container_width=True, type="primary"):
                st.session_state["enviado_exitoso"] = False
                st.rerun()
        with c_btn2:
            if st.button("🏠 Volver al inicio", use_container_width=True):
                st.session_state["enviado_exitoso"] = False
                st.session_state["form_logueado"] = False
                st.session_state["nav_state"] = "home"
                st.rerun()
    else:
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
                total_fruta = st.number_input("Total de Fruta Ingresada", min_value=0.0, value=0.0)
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
                        "Responsable": responsable, "Fecha": str(fecha), "Hora": str(hora),
                        "Desc_Materia": desc_materia, "Observaciones": observaciones,
                        "Proveedor": proveedor_a_guardar, "Total_Fruta": total_fruta, "Cant_Unidades": cant_unidades,
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
        st.markdown("Ingrese la contraseña de administrador (`glad726lif`):")
        
        password_input = st.text_input("Contraseña de administrador", type="password")
        if st.button("Verificar Acceso", use_container_width=True, type="primary"):
            if password_input == "glad726lif":  
                st.session_state["admin_logueado"] = True
                st.session_state["nav_state"] = "admin_dashboard"
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")

# ==========================================
# 5. DASHBOARD DEL ADMINISTRADOR
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
        tab_pendientes, tab_aprobados, tab_todos = st.tabs([
            "⏳ Registros Pendientes por Validar", 
            "✅ Registros Aprobados", 
            "📊 Historial Completo de Registros"
        ])
        
        # PESTAÑA 1: PENDIENTES
        with tab_pendientes:
            df_pendientes = df[df["Estado"] == "Pendiente"]
            if df_pendientes.empty:
                st.success("¡Excelente! No hay registros pendientes de firma.")
            else:
                for idx, row in df_pendientes.iterrows():
                    st.markdown(f"""
                    <div class="card-seccion">
                    <b>Registrado el {row['Fecha']} por {row['Responsable']} · Estado: Pendiente de revisión</b><br><br>
                    <b>📦 Datos Generales:</b><br>
                    Proveedor: {row['Proveedor']} | Materia Prima: {row['Desc_Materia']} | Total Fruta: {row['Total_Fruta']} L<br>
                    Hora: {row['Hora']} | Observaciones: {row['Observaciones']}
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Botón para eliminar registro de prueba
                    if st.button(f"🗑️ Eliminar Registro #{row['ID_Registro']}", key=f"del_pen_{row['ID_Registro']}"):
                        eliminar_registro(row['ID_Registro'])
                        st.rerun()
                    
                    with st.expander(f"🔍 Revisar y completar aprobación (ID: #{row['ID_Registro']})"):
                        st.markdown("#### Parámetros Registrados:")
                        c_m1, c_m2, c_m3 = st.columns(3)
                        with c_m1:
                            st.markdown(f"**Muestra 1**<br>Galón: {row['unidades_galon_1']}<br>Vol: {row['volumen_1']}<br>Brix: {row['brix_1']}<br>pH: {row['ph_1']}", unsafe_allow_html=True)
                        with c_m2:
                            st.markdown(f"**Muestra 2**<br>Galón: {row['unidades_galon_2']}<br>Vol: {row['volumen_2']}<br>Brix: {row['brix_2']}<br>pH: {row['ph_2']}", unsafe_allow_html=True)
                        with c_m3:
                            st.markdown(f"**Muestra 3**<br>Galón: {row['unidades_galon_3']}<br>Vol: {row['volumen_3']}<br>Brix: {row['brix_3']}<br>pH: {row['ph_3']}", unsafe_allow_html=True)
                        
                        st.markdown("---")
                        st.markdown("#### ✍️ Firma de V°B° Calidad")
                        
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
                        
                        if st.button(f"Guardar aprobación y firmar #{row['ID_Registro']}", key=f"btn_firmar_{row['ID_Registro']}", type="primary"):
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
                    st.markdown("---")

        # PESTAÑA 2: APROBADOS
        with tab_aprobados:
            df_aprobados = df[df["Estado"] == "Aprobado"]
            if df_aprobados.empty:
                st.info("Aún no hay registros aprobados.")
            else:
                for idx, row in df_aprobados.iterrows():
                    col_info, col_del = st.columns([5, 1])
                    with col_info:
                        st.markdown(f"""
                        <div class="card-seccion">
                        <b>✅ ID: {row['ID_Registro']} | Proveedor: {row['Proveedor']}</b><br>
                        Fecha: {row['Fecha']} | Responsable: {row['Responsable']} | Estado: Aprobado
                        </div>
                        """, unsafe_allow_html=True)
                    with col_del:
                        if st.button("🗑️", key=f"del_apr_{row['ID_Registro']}", help="Eliminar registro"):
                            eliminar_registro(row['ID_Registro'])
                            st.rerun()

        # PESTAÑA 3: HISTORIAL COMPLETO
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
                    <div class="card-seccion">
                    <b>ID: {row['ID_Registro']}</b> | <b>Fecha:</b> {row['Fecha']} | <b>Proveedor:</b> {row['Proveedor']} | <b>Estado:</b> {estado_badge}
                    </div>
                    """, unsafe_allow_html=True)
                    
                    c_btn1, c_btn2 = st.columns([2, 8])
                    
                    with c_btn1:
                        if st.button("🗑️ Eliminar", key=f"del_tod_{row['ID_Registro']}"):
                            eliminar_registro(row['ID_Registro'])
                            st.rerun()
                    
                    with c_btn2:
                        if row["Estado"] == "Aprobado":
                            pdf_bytes = generar_pdf_nuevo(row.to_dict())
                            st.download_button(
                                label=f"📥 Descargar PDF Relleno y Firmado",
                                data=pdf_bytes,
                                file_name=f"Control_Coco_{row['ID_Registro']}.pdf",
                                mime="application/pdf",
                                key=f"download_{row['ID_Registro']}"
                            )
                    st.markdown("---")
