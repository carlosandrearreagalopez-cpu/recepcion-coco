import streamlit as st
import pandas as pd
import os
from datetime import datetime
from streamlit_drawable_canvas import st_canvas
import io

# Importaciones para Excel
import openpyxl

# Importaciones para PDF (Formato Horizontal Matricial)
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.platypus import Table, TableStyle, Image as RLImage
from reportlab.lib import colors

# ==========================================
# CONFIGURACIÓN Y ESTILOS (Colores LIF Brands)
# ==========================================
st.set_page_config(page_title="Control de Recepción - LIF Brands", layout="wide")

st.markdown("""
<style>
/* Forzar fondo blanco y texto oscuro para garantizar legibilidad */
.stApp { background-color: #F8FAF9 !important; }
html, body, [class*="css"], p, span, label, div { 
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif !important; 
    color: #1f2937 !important; 
}
h1, h2, h3, h4, h5, h6 { color: #115e59 !important; font-weight: 700 !important; }

/* Estilos de Inputs y Selects */
.stTextInput label, .stSelectbox label, .stDateInput label, .stNumberInput label, .stRadio label, .stFileUploader label {
    color: #0f766e !important;
    font-weight: bold !important;
}
input, select, textarea {
    background-color: #ffffff !important;
    color: #1f2937 !important;
    border: 1px solid #cbd5e1 !important;
    border-radius: 6px;
}

/* Botones Principales (Verde LIF) */
button[kind="primary"] {
    background-color: #115e59 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: bold !important;
}
button[kind="primary"]:hover {
    background-color: #0f766e !important;
}

/* Botones Secundarios */
.stButton>button {
    background-color: #FFFFFF !important;
    color: #115e59 !important;
    border: 1px solid #115e59 !important;
    border-radius: 6px !important;
    font-weight: bold;
}
.stButton>button:hover {
    background-color: #f0fdf4 !important;
}

/* Tarjetas (Cards) de Registros imitando las imágenes */
.record-card {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 16px;
    margin-bottom: 12px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    border-left: 5px solid #0f766e;
}
.record-header {
    font-size: 1.1em;
    font-weight: bold;
    color: #1f2937;
    margin-bottom: 8px;
}
.record-sub {
    color: #64748b;
    font-size: 0.9em;
}

/* Expanders (Acordeones de Streamlit) */
.streamlit-expanderHeader {
    color: #0f766e !important;
    background-color: #f8fafc !important;
    border-radius: 4px;
    font-weight: bold !important;
}
.streamlit-expanderContent {
    background-color: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-top: none !important;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# DIRECTORIOS Y FUNCIONES BASE
# ==========================================
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
    st.success(f"Registro #{id_registro} eliminado.")

def generar_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Registros Completos')
    return output.getvalue()

# ==========================================
# GENERADOR DE PDF (Formato Horizontal)
# ==========================================
def generar_pdf_nuevo(registro):
    buffer = io.BytesIO()
    c = pdf_canvas.Canvas(buffer, pagesize=landscape(letter))
    width, height = landscape(letter)
    
    margin_x, margin_y = 40, 40
    usable_width = width - (2 * margin_x)
    col_widths = [160, 220, 60, 100, 50, usable_width - (160+220+60+100+50)]
    
    logo_img = RLImage("logo.png", width=90, height=35) if os.path.exists("logo.png") else ""
        
    data = [
        [logo_img, "Control de Recepción de Coco", "", "", "", f"Código: R ICC/15-2\nVersión: 02\nID: #{registro.get('ID_Registro', '')}"],
        ["Nombre del responsable:", registro.get('Responsable', ''), "Fecha:", str(registro.get('Fecha', '')), "Hora:", str(registro.get('Hora', ''))],
        ["Descripción de materia prima:", registro.get('Desc_Materia', ''), "Observaciones", "", "", ""],
        ["Nombre del proveedor", registro.get('Proveedor', ''), registro.get('Observaciones', 'Ninguna'), "", "", ""],
        ["Total de Fruta Ingresada Planta", str(registro.get('Total_Fruta', '')), "", "", "", ""],
        ["Cantidad Unidades (Muestra)", str(registro.get('Cant_Unidades', '')), "", "", "", ""],
        ["PARÁMETROS FISICOQUÍMICOS", "", "", "", "", ""],
        ["Muestra 1", "", "", "", "", ""],
        ["Unidades/Galón", str(registro.get('unidades_galon_1', '')), "", "", "", ""],
        ["Volumen (Muestra)", str(registro.get('volumen_1', '')), "", "", "", ""],
        ["Brix° (5 - 5.9)", str(registro.get('brix_1', '')), "", "", "", ""],
        ["pH", str(registro.get('ph_1', '')), "", "", "", ""],
        ["Muestra 2", "", "", "", "", ""],
        ["Unidades/Galón", str(registro.get('unidades_galon_2', '')), "", "", "", ""],
        ["Volumen (Muestra)", str(registro.get('volumen_2', '')), "", "", "", ""],
        ["Brix° (5 - 5.9)", str(registro.get('brix_2', '')), "", "", "", ""],
        ["pH", str(registro.get('ph_2', '')), "", "", "", ""],
        ["Muestra 3", "", "", "", "", ""],
        ["Unidades/Galón", str(registro.get('unidades_galon_3', '')), "", "", "", ""],
        ["Volumen (Muestra)", str(registro.get('volumen_3', '')), "", "", "", ""],
        ["Brix° (5 - 5.9)", str(registro.get('brix_3', '')), "", "", "", ""],
        ["pH", str(registro.get('ph_3', '')), "", "", "", ""],
    ]
    
    style = TableStyle([
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('BOX', (0,0), (-1,-1), 2, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('SPAN', (1,0), (4,0)), ('ALIGN', (1,0), (1,0), 'CENTER'), ('FONTNAME', (1,0), (1,0), 'Helvetica-Bold'), ('FONTSIZE', (1,0), (1,0), 14),
        ('FONTSIZE', (5,0), (5,0), 8), ('VALIGN', (5,0), (5,0), 'TOP'),
        ('SPAN', (2,2), (5,2)), ('ALIGN', (2,2), (5,2), 'CENTER'), ('FONTNAME', (2,2), (5,2), 'Helvetica-Bold'),
        ('SPAN', (2,3), (5,5)), ('VALIGN', (2,3), (5,5), 'TOP'),
        ('SPAN', (1,3), (1,3)), ('SPAN', (1,4), (1,4)), ('SPAN', (1,5), (1,5)),
        ('SPAN', (0,6), (-1,6)), ('ALIGN', (0,6), (-1,6), 'CENTER'), ('FONTNAME', (0,6), (-1,6), 'Helvetica-Bold'),
        ('SPAN', (0,7), (-1,7)), ('ALIGN', (0,7), (-1,7), 'CENTER'), ('FONTNAME', (0,7), (-1,7), 'Helvetica-Bold'),
        ('SPAN', (1,8), (-1,8)), ('SPAN', (1,9), (-1,9)), ('SPAN', (1,10), (-1,10)), ('SPAN', (1,11), (-1,11)),
        ('SPAN', (0,12), (-1,12)), ('ALIGN', (0,12), (-1,12), 'CENTER'), ('FONTNAME', (0,12), (-1,12), 'Helvetica-Bold'),
        ('SPAN', (1,13), (-1,13)), ('SPAN', (1,14), (-1,14)), ('SPAN', (1,15), (-1,15)), ('SPAN', (1,16), (-1,16)),
        ('SPAN', (0,17), (-1,17)), ('ALIGN', (0,17), (-1,17), 'CENTER'), ('FONTNAME', (0,17), (-1,17), 'Helvetica-Bold'),
        ('SPAN', (1,18), (-1,18)), ('SPAN', (1,19), (-1,19)), ('SPAN', (1,20), (-1,20)), ('SPAN', (1,21), (-1,21)),
    ])
    
    # === AQUÍ ESTÁ LA CORRECCIÓN DE LAS ALTURAS (22 Filas en total) ===
    row_heights = [45] + [20]*5 + [15]*16
    t = Table(data, colWidths=col_widths, rowHeights=row_heights)
    t.setStyle(style)
    
    w, h = t.wrapOn(c, usable_width, height)
    y_pos_table = height - margin_y - h
    t.drawOn(c, margin_x, y_pos_table)
    
    y_firma = y_pos_table - 45
    nombre_firma = registro.get("Firma_Jefe", "Sin firma")
    ruta_firma = os.path.join(FIRMAS_DIR, str(nombre_firma))
    
    if nombre_firma != "Sin firma" and os.path.exists(ruta_firma):
        c.drawImage(ruta_firma, width/2.0 - 60, y_firma, width=120, height=40, preserveAspectRatio=True, mask='auto')
        
    c.line(width/2.0 - 120, y_firma - 5, width/2.0 + 120, y_firma - 5)
    c.setFont("Helvetica-Bold", 9)
    c.drawCentredString(width/2.0, y_firma - 18, "Jefe de Calidad")
    
    c.save()
    buffer.seek(0)
    return buffer.getvalue()

# ==========================================
# CONTROL DE ESTADOS DE SESIÓN
# ==========================================
for state in ["nav_state", "form_logueado", "admin_logueado", "enviado_exitoso"]:
    if state not in st.session_state:
        st.session_state[state] = "home" if state == "nav_state" else False

# ==========================================
# 1. PANTALLA DE INICIO
# ==========================================
if st.session_state["nav_state"] == "home":
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        mostrar_logo(200)
        st.markdown("<h1 style='text-align: center; color: #115e59 !important;'>Control de Recepción de Coco</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; font-weight: bold;'>LIF Brands - Aseguramiento de Calidad</p>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("📝 Colaborador: Reportar nuevo ingreso", use_container_width=True, type="primary"):
            st.session_state["nav_state"] = "form_login"
            st.rerun()
            
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔒 Jefe de Calidad: Panel de Administración", use_container_width=True):
            st.session_state["nav_state"] = "admin_login"
            st.rerun()

# ==========================================
# 2. LOGIN (AMBOS PERFILES)
# ==========================================
elif st.session_state["nav_state"] in ["form_login", "admin_login"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("⬅️ Volver al inicio"):
            st.session_state["nav_state"] = "home"
            st.rerun()
        
        es_admin = st.session_state["nav_state"] == "admin_login"
        st.title("Panel de Administrador" if es_admin else "Acceso a Registro")
        st.markdown("Ingrese la contraseña autorizada:")
        
        pwd = st.text_input("Contraseña", type="password")
        pwd_correcta = "glad726lif" if es_admin else "20lf26"
        
        if st.button("Verificar Acceso", use_container_width=True, type="primary"):
            if pwd == pwd_correcta:
                st.session_state["admin_logueado" if es_admin else "form_logueado"] = True
                st.session_state["nav_state"] = "admin_dashboard" if es_admin else "form"
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")

# ==========================================
# 3. FORMULARIO DEL COLABORADOR
# ==========================================
elif st.session_state["nav_state"] == "form":
    if not st.session_state["form_logueado"]:
        st.session_state["nav_state"] = "home"
        st.rerun()
        
    if st.button("⬅️ Cerrar Sesión"):
        st.session_state["form_logueado"] = False
        st.session_state["nav_state"] = "home"
        st.rerun()
    
    mostrar_logo(140)
    st.title("Registro de Recepción de Coco")
    
    if st.session_state["enviado_exitoso"]:
        st.success("¡Registro enviado con éxito! Quedará pendiente de validación en el panel de Jefe de Calidad.")
        if st.button("➕ Ingresar un nuevo registro", type="primary"):
            st.session_state["enviado_exitoso"] = False
            st.rerun()
    else:
        with st.form("form_coco"):
            st.header("1. Datos Generales")
            c1, c2 = st.columns(2)
            with c1:
                responsable = st.selectbox("Nombre del responsable", ["Carlos Canto", "Carlos Rodas", "Jonathan", "Damarias Arellanos", "Carlos López", "Marlon Escobar"])
                proveedor_opcion = st.selectbox("Proveedor", ["GRANOS BASICOS LA PATRONA SOCIEDAD ANONIMA", "Otro"])
                proveedor_final = st.text_input("Si eligió 'Otro', escriba el nombre:") if proveedor_opcion == "Otro" else proveedor_opcion
                desc_materia = st.text_input("Materia prima", value="Coco")
            with c2:
                fecha = st.date_input("Fecha")
                hora = st.time_input("Hora")
                total_fruta = st.number_input("Total Fruta Ingresada", min_value=0.0)
                cant_unidades = st.number_input("Unidades (Muestra)", min_value=0.0)
            
            observaciones = st.text_area("Observaciones", value="Ninguna")
            
            st.header("2. Parámetros Fisicoquímicos")
            muestras_datos = {}
            for i in range(1, 4):
                st.subheader(f"Muestra {i}")
                mc1, mc2, mc3, mc4 = st.columns(4)
                with mc1: muestras_datos[f"ug_{i}"] = st.number_input(f"Unidades/Galón (M{i})", min_value=0.0)
                with mc2: muestras_datos[f"v_{i}"] = st.number_input(f"Volumen (M{i})", min_value=0.0)
                with mc3: muestras_datos[f"b_{i}"] = st.number_input(f"Brix° (M{i})", min_value=0.0, format="%.2f")
                with mc4: muestras_datos[f"ph_{i}"] = st.number_input(f"pH (M{i})", min_value=0.0, format="%.2f")

            submitted = st.form_submit_button("Guardar y Enviar a Revisión", type="primary")
            
            if submitted:
                prov = proveedor_final if proveedor_opcion == "Otro" else proveedor_opcion
                nuevo_registro = {
                    "ID_Registro": datetime.now().strftime("%Y%m%d%H%M%S"),
                    "Estado": "Pendiente",
                    "Responsable": responsable, "Fecha": str(fecha), "Hora": str(hora),
                    "Desc_Materia": desc_materia, "Observaciones": observaciones,
                    "Proveedor": prov, "Total_Fruta": total_fruta, "Cant_Unidades": cant_unidades,
                    "unidades_galon_1": muestras_datos["ug_1"], "volumen_1": muestras_datos["v_1"], "brix_1": muestras_datos["b_1"], "ph_1": muestras_datos["ph_1"],
                    "unidades_galon_2": muestras_datos["ug_2"], "volumen_2": muestras_datos["v_2"], "brix_2": muestras_datos["b_2"], "ph_2": muestras_datos["ph_2"],
                    "unidades_galon_3": muestras_datos["ug_3"], "volumen_3": muestras_datos["v_3"], "brix_3": muestras_datos["b_3"], "ph_3": muestras_datos["ph_3"],
                    "Firma_Jefe": "Sin firma"
                }
                df = cargar_datos()
                df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
                guardar_datos(df)
                st.session_state["enviado_exitoso"] = True
                st.rerun()

# ==========================================
# 4. DASHBOARD DEL ADMINISTRADOR (Estilo LIF)
# ==========================================
elif st.session_state["nav_state"] == "admin_dashboard":
    if not st.session_state.get("admin_logueado", False):
        st.session_state["nav_state"] = "home"
        st.rerun()
        
    c_head1, c_head2 = st.columns([5, 1])
    with c_head1:
        st.title("Panel de Administrador - Jefe de Calidad")
        st.markdown("Sesión: Administrador (Calidad LIF Brands)")
    with c_head2:
        if st.button("Cerrar sesión"):
            st.session_state["admin_logueado"] = False
            st.session_state["nav_state"] = "home"
            st.rerun()
            
    df = cargar_datos()
    
    # Diseño de pestañas tipo interfaz mostrada
    tab_pendientes, tab_aprobados, tab_todos = st.tabs([
        "⏳ Pendientes", "✅ Aprobados", "📊 Todos"
    ])
    
    # FUNCIONES DE RENDERIZADO DE TARJETAS
    def render_tarjeta(row, index_key):
        estado_icono = "⏳" if row['Estado'] == "Pendiente" else "✅"
        st.markdown(f"""
        <div class="record-card">
            <div class="record-header">{estado_icono} #{row['ID_Registro']} — {row['Proveedor']} · {row['Desc_Materia']}</div>
            <div class="record-sub">Bodega Materia Prima | Fecha: {row['Fecha']} | Receptor: {row['Responsable']} | Estado: <b>{row['Estado']}</b></div>
        </div>
        """, unsafe_allow_html=True)
        
        c_btn1, c_btn2, c_btn3 = st.columns([2, 2, 8])
        
        with c_btn1:
            if st.button("🗑️ Eliminar", key=f"del_{index_key}_{row['ID_Registro']}", help="Eliminar registro"):
                eliminar_registro(row['ID_Registro'])
                st.rerun()
                
        with c_btn2:
            if row['Estado'] == 'Aprobado':
                pdf_bytes = generar_pdf_nuevo(row.to_dict())
                st.download_button("📥 PDF", data=pdf_bytes, file_name=f"Registro_{row['ID_Registro']}.pdf", mime="application/pdf", key=f"pdf_{index_key}_{row['ID_Registro']}")

        with st.expander(f"Ver detalles del registro #{row['ID_Registro']}"):
            st.write(f"**Observaciones:** {row['Observaciones']} | **Fruta:** {row['Total_Fruta']} | **Unidades:** {row['Cant_Unidades']}")
            st.write("---")
            c_m1, c_m2, c_m3 = st.columns(3)
            with c_m1: st.info(f"**M1:** Galón: {row['unidades_galon_1']} | Vol: {row['volumen_1']} | Brix: {row['brix_1']} | pH: {row['ph_1']}")
            with c_m2: st.info(f"**M2:** Galón: {row['unidades_galon_2']} | Vol: {row['volumen_2']} | Brix: {row['brix_2']} | pH: {row['ph_2']}")
            with c_m3: st.info(f"**M3:** Galón: {row['unidades_galon_3']} | Vol: {row['volumen_3']} | Brix: {row['brix_3']} | pH: {row['ph_3']}")
            
            if row['Estado'] == "Pendiente":
                st.markdown("#### ✍️ Aprobar y Firmar")
                canvas_result = st_canvas(
                    fill_color="rgba(255, 255, 255, 0)", stroke_width=2, stroke_color="#0f766e",
                    background_color="#f8fafc", height=100, width=350, drawing_mode="freedraw",
                    key=f"canvas_{row['ID_Registro']}"
                )
                if st.button("Aprobar Registro", key=f"btn_aprobar_{row['ID_Registro']}", type="primary"):
                    if canvas_result.image_data is not None:
                        from PIL import Image
                        img = Image.fromarray(canvas_result.image_data.astype('uint8'), mode="RGBA")
                        nombre_firma = f"firma_{row['ID_Registro']}.png"
                        img.save(os.path.join(FIRMAS_DIR, nombre_firma))
                        
                        df.loc[df['ID_Registro'] == row['ID_Registro'], 'Estado'] = "Aprobado"
                        df.loc[df['ID_Registro'] == row['ID_Registro'], 'Firma_Jefe'] = nombre_firma
                        guardar_datos(df)
                        st.success("Registro aprobado correctamente.")
                        st.rerun()

    # --- PESTAÑA: PENDIENTES ---
    with tab_pendientes:
        df_pen = df[df["Estado"] == "Pendiente"]
        st.write(f"### Registros Pendientes ({len(df_pen)})")
        for idx, row in df_pen.iterrows():
            render_tarjeta(row, "pen")

    # --- PESTAÑA: APROBADOS ---
    with tab_aprobados:
        df_apr = df[df["Estado"] == "Aprobado"]
        st.write(f"### Registros Aprobados ({len(df_apr)})")
        for idx, row in df_apr.iterrows():
            render_tarjeta(row, "apr")

    # --- PESTAÑA: TODOS Y DESCARGA EXCEL ---
    with tab_todos:
        st.write("### Historial Completo")
        
        # Panel de Métricas
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        col_m1.metric("⏳ Pendientes", len(df[df["Estado"]=="Pendiente"]))
        col_m2.metric("✅ Aprobados", len(df[df["Estado"]=="Aprobado"]))
        col_m3.metric("❌ Rechazados", 0) # Placeholder si en el futuro hay rechazados
        col_m4.metric("📊 Total", len(df))
        
        st.markdown("---")
        
        # Botón Descargar Excel Integrado
        if not df.empty:
            excel_bytes = generar_excel_bytes(df)
            st.download_button(
                label=f"📥 Descargar Excel completo ({len(df)} filas)",
                data=excel_bytes,
                file_name="Todos_los_Registros_LIF.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Filtros opcionales estilo LIF
        cf1, cf2 = st.columns(2)
        with cf1:
            prov_filtro = st.selectbox("Filtrar por Proveedor:", ["Todos"] + list(df["Proveedor"].unique()))
        
        df_mostrar = df if prov_filtro == "Todos" else df[df["Proveedor"] == prov_filtro]
        
        for idx, row in df_mostrar.iterrows():
            render_tarjeta(row, "tod")
