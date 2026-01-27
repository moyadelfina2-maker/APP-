"""
app_streamlit.py
Aplicación Streamlit que reproduce la interfaz original en Colab/Jupyter:
- Todos los campos principales están incluidos
- Lógica dinámica para añadir/eliminar hallazgos y filas de muestreo
- Usa docgen.generate_report(...) para la generación del .docx

Ejecutar:
    streamlit run app_streamlit.py
"""
import streamlit as st
import docgen
import io
from datetime import datetime

st.set_page_config(page_title="Generador de Informes - APP", layout="wide")

# -----------------------
# CONSTANTES / OPCIONES (coinciden con las del módulo original)
# -----------------------
CAAP_LOGICA_ESTADOS = [
    "No aplica",
    "No iniciado, tiene CNCA vigente",
    "No iniciado, tiene CNCA en curso",
    "No iniciado, tiene CNCA vencida / No solicitada)",
    "En curso",
    "Vigente",
    "Vencido"
]
ADA_DETALLE_ESTADOS = ["Seleccione...", "Prefactibilidad con Chi 0 para todos los permisos", "Prefactibilidad vigente", "Prefactibilidad vencida", "No solicitada"]
RENPRE_DETALLE_ESTADOS = ["Seleccione...", "No aplica", "Esta inscripto y renueva", "Esta inscripto y no renovó", "Aplica pero no esta inscripto"]
RESIDUOS_ESPECIALES_STATUS_GENERAL = [
    "Seleccione...", "Empresa exenta porque no genera", "No inscripta", "No Cumple con las DDJJ",
    "Cumple con las DDJJ"
]
CHE_DETALLE_ESTADOS = [
    "Seleccione...", "Obtuvo CHE", "No obtuvo CHE", "CHE en curso"
]
FRECUENCIA_OPTIONS = ['Seleccione...', 'Mensual', 'Trimestral', 'Semestral', 'Anual', 'N/A']
ACUMAR_DETALLE_ESTADOS = ["Seleccione...", "No aplica", "Vigente", "En curso", "Vencida", "No solicitada"]
SE_DETALLE_ESTADOS = ["Seleccione...", "No aplica", "Vigente", "En curso", "Vencida", "No solicitada"]

# -----------------------
# Helpers de sesión
# -----------------------
if "hallazgos" not in st.session_state:
    st.session_state.hallazgos = []
if "muestreo_rows" not in st.session_state:
    st.session_state.muestreo_rows = []

# -----------------------
# UI: encabezado y subida de plantilla
# -----------------------
st.title("📝 Generador de Informes Ambientales (Streamlit)")
st.write("Sube la plantilla DOCX (con marcadores tipo {RAZON_SOCIAL}) y completa los campos. Luego presiona GENERAR INFORME.")

template_file = st.file_uploader("Cargar plantilla .docx", type=["docx"]) 

# -----------------------
# FORMULARIO PRINCIPAL (campos replicados)
# -----------------------
with st.form("form_principal"):
    st.header("1. Información General")
    col1, col2 = st.columns(2)
    with col1:
        razon_social = st.text_input("Razón Social", value="EMPRESA S.A.")
        nombre_planta = st.text_input("Planta Industrial", value="Planta Industrial")
        mes_auditoria = st.text_input("Mes de relevamiento (Mes/Año)", value=datetime.now().strftime("%B %Y"))
        direccion = st.text_input("Dirección", value="")
    with col2:
        municipio = st.text_input("Municipio", value="")
        rubro = st.text_input("Rubro", value="")
        # Habilitación municipal
        st.subheader("Habilitación municipal")
        hab_status = st.selectbox("Estado Habilitación", options=['Seleccione...', 'cumple', 'no cumple', 'parcial'])
        hab_fecha = st.text_input("Fecha de obtención de habilitación", value="dd/mm/aaaa")
        hab_expediente = st.text_input("Nº de expediente (Habilitación)", value="N/A")
        hab_obs = st.text_area("Observaciones habilitación municipal", value="")

    st.markdown("---")
    st.header("2. Permisos Centrales (CNCA, CAAP, CAAF)")
    col3, col4 = st.columns(2)
    with col3:
        cnca_status = st.selectbox("Estado CNCA", options=["No aplica", "vigente", "superada (esta en curso el CAA)", "vencida", "no solicitada"])
        cnca_fecha = st.text_input("Fecha CNCA", value="dd/mm/aaaa")
        cnca_vto = st.text_input("Vencimiento CNCA", value="dd/mm/aaaa")
        cnca_expediente = st.text_input("Expediente CNCA", value="N/A")
        cnca_categoria = st.text_input("Categoría CNCA", value="N/A")
        cnca_puntos = st.text_input("Puntos CNCA", value="")
    with col4:
        caap_status = st.selectbox("Situación CAAP", options=CAAP_LOGICA_ESTADOS)
        caap_fecha = st.text_input("Fecha obtención CAAP", value="N/A")
        caap_expediente = st.text_input("Expediente CAAP", value="N/A")
        caaf_status = st.selectbox("Situación CAAF", options=["No aplica", "no_iniciado_caap_en_curso", "no_iniciado_caap_vencido", "en_curso", "vigente", "vencido"])
        caaf_fecha = st.text_input("Fecha obtención CAAF", value="N/A")
        renovacion_caa_status = st.selectbox("Renovación CAA", options=["No aplica", "En curso", "Finalizada", "No iniciada"])

    st.markdown("---")
    st.header("3. LEGA y Plan de Monitoreos")
    col5, col6 = st.columns(2)
    with col5:
        lega_status = st.selectbox("Estado LEGA", options=["Seleccione...", "vigente", "en_curso", "vencida"])
        lega_fecha = st.text_input("Fecha obtención LEGA", value="N/A")
        lega_expediente = st.text_input("Expediente LEGA", value="N/A")
        lega_estado_portal = st.text_input("Estado LEGA en portal", value="N/A")
    with col6:
        # Muestreo: mostrador sencillo; filas dinámicas más abajo
        st.info("Añade filas al Plan de Monitoreos más abajo")

    st.markdown("---")
    st.header("4. Autoridad del Agua (ADA) y Otros")
    col7, col8 = st.columns(2)
    with col7:
        ada_status = st.selectbox("Estado ADA", options=ADA_DETALLE_ESTADOS)
        ada_fecha = st.text_input("Fecha Prefactibilidad", value="N/A")
        ada_expediente = st.text_input("Expediente Prefactibilidad", value="N/A")
        chi_hid = st.text_input("CHi Hidráulica", value="0")
    with col8:
        chi_exp = st.text_input("CHi Explotación", value="0")
        chi_vue = st.text_input("CHi Vuelco", value="0")
        red_monitoreos = st.text_input("Red de Monitoreos", value="N/A")
        renpre_status = st.selectbox("Estado RENPRE", options=RENPRE_DETALLE_ESTADOS)

    st.markdown("---")
    st.header("5. Residuos Especiales y CHE")
    col9, col10 = st.columns(2)
    with col9:
        rree_status = st.selectbox("Estado Generador (RREE)", options=RESIDUOS_ESPECIALES_STATUS_GENERAL)
        che_status = st.selectbox("Estado CHE", options=CHE_DETALLE_ESTADOS)
        anio_che = st.text_input("Año CHE", value=str(datetime.now().year))
    with col10:
        obs_che = st.text_area("Observaciones CHE", value="")
        gestion_res = st.selectbox("Gestión operativa residuos", options=["Correcta", "Mala"])
        tipo_res = st.text_input("Tipo de residuo", value="")

    st.markdown("---")
    st.header("6. ASP (Aparatos a Presión)")
    asp_status = st.selectbox("Estado ASP", options=["Seleccione...", "Finalizada", "Caratulada", "No Presentado"])
    vto_asp = st.text_input("Vencimiento ASP", value="N/A")
    expediente_asp = st.text_input("Expediente ASP", value="N/A")
    valvulas_status = st.selectbox("Calibración válvulas", options=["Cumple","No Cumple"])
    vto_valvulas = st.text_input("Vto Calibración válvulas", value="N/A")

    # Submit del formulario principal (no genera aún)
    form_submitted = st.form_submit_button("Guardar datos (no genera aún)")

# -----------------------
# Sección muestreo dinámico
# -----------------------
st.header("Plan de Monitoreos - Filas dinámicas")
col_a, col_b = st.columns([3,1])
with col_a:
    for i, fila in enumerate(st.session_state.muestreo_rows):
        cols = st.columns([3,3,2,3,2])
        with cols[0]:
            recurso = st.text_input(f"Recurso #{i+1}", value=fila.get('recurso',''), key=f"rec_{i}")
        with cols[1]:
            organismo = st.text_input(f"Organismo #{i+1}", value=fila.get('organismo',''), key=f"org_{i}")
        with cols[2]:
            puntos = st.text_input(f"Puntos #{i+1}", value=fila.get('puntos',''), key=f"pun_{i}")
        with cols[3]:
            parametros = st.text_input(f"Parámetros #{i+1}", value=fila.get('parametros',''), key=f"par_{i}")
        with cols[4]:
            frecuencia = st.selectbox(f"Frecuencia #{i+1}", options=FRECUENCIA_OPTIONS, index=FRECUENCIA_OPTIONS.index(fila.get('frecuencia','Seleccione...')) if fila.get('frecuencia') in FRECUENCIA_OPTIONS else 0, key=f"frec_{i}")
        # Actualizar estado con los valores retornados (reconstruimos la fila)
        st.session_state.muestreo_rows[i] = {
            'recurso': st.session_state.get(f"rec_{i}", fila.get('recurso','')),
            'organismo': st.session_state.get(f"org_{i}", fila.get('organismo','')),
            'puntos': st.session_state.get(f"pun_{i}", fila.get('puntos','')),
            'parametros': st.session_state.get(f"par_{i}", fila.get('parametros','')),
            'frecuencia': st.session_state.get(f"frec_{i}", fila.get('frecuencia','Seleccione...'))
        }
with col_b:
    if st.button("Añadir fila"):
        st.session_state.muestreo_rows.append({'recurso':'','organismo':'','puntos':'','parametros':'','frecuencia':'Seleccione...'})
        st.experimental_rerun()
    if st.session_state.muestreo_rows:
        if st.button("Eliminar última fila"):
            st.session_state.muestreo_rows.pop()
            st.experimental_rerun()

# -----------------------
# Hallazgos dinámicos (añadir/eliminar)
# -----------------------
st.header("Hallazgos de Campo")
with st.expander("Añadir nuevo hallazgo"):
    h_obs = st.text_area("Observación", key="new_h_obs")
    h_sit = st.text_area("Situación", key="new_h_sit")
    h_aut = st.text_input("Autoridad", key="new_h_aut")
    h_rie = st.text_input("Riesgo", key="new_h_rie")
    h_rec = st.text_area("Recomendación", key="new_h_rec")
    if st.button("Añadir hallazgo"):
        st.session_state.hallazgos.append({
            'observacion': h_obs,
            'situacion': h_sit,
            'autoridad': h_aut,
            'riesgo': h_rie,
            'recomendacion': h_rec
        })
        st.experimental_rerun()

for idx, h in enumerate(st.session_state.hallazgos):
    with st.expander(f"Observación #{idx+1}", expanded=False):
        st.write("Observación:", h.get('observacion',''))
        st.write("Situación:", h.get('situacion',''))
        st.write("Autoridad:", h.get('autoridad',''))
        st.write("Riesgo:", h.get('riesgo',''))
        st.write("Recomendación:", h.get('recomendacion',''))
        if st.button(f"Eliminar #{idx+1}", key=f"del_h_{idx}"):
            st.session_state.hallazgos.pop(idx)
            st.experimental_rerun()

# -----------------------
# Botón GENERAR INFORME (usa docgen)
# -----------------------
st.markdown("---")
if st.button("GENERAR INFORME"):
    if not template_file:
        st.error("Sube primero una plantilla .docx")
    else:
        # Construir user_data con todas las claves necesarias (nombres en mayúscula para el reemplazo)
        user_data = {
            'RAZON_SOCIAL': razon_social,
            'NOMBRE_EMPRESA': razon_social,
            'NOMBRE_PLANTA': nombre_planta,
            'MES_AUDITORIA': mes_auditoria,
            'DIRECCION_EMPRESA': direccion,
            'MUNICIPIO_EMPRESA': municipio,
            'RUBRO_EMPRESA': rubro,
            'HABILITACION_MUNICIPAL_STATUS': hab_status,
            'FECHA_HABILITACIÓN': hab_fecha,
            'EXPEDIENTE_HABILITACION': hab_expediente,
            'OBSERVACION_HAB_MUNICIPAL': hab_obs,
            'CNCA_DETALLE_STATUS': cnca_status,
            'FECHA_CNCA': cnca_fecha,
            'FECHA_CNCA_VENCIMIENTO': cnca_vto,
            'EXPEDIENTE_CNCA': cnca_expediente,
            'CATEGORIA_CNCA': cnca_categoria,
            'PUNTOS_CNCA': cnca_puntos,
            'CAAP_LOGICA': caap_status,
            'FECHA_Obtención_CAAP': caap_fecha,
            'EXPEDIENTE_CAAP': caap_expediente,
            'CAAF_STATUS': caaf_status,
            'RENOVACION_CAA_STATUS': renovacion_caa_status,
            'LEGA_STATUS': lega_status,
            'FECHA_OBTENCION_LEGA': lega_fecha,
            'EXPEDIENTE_LEGA': lega_expediente,
            'ESTADO_LEGA': lega_estado_portal,
            'ADA_STATUS': ada_status,
            'FECHA_PREFA': ada_fecha,
            'EXPEDIENTE_PREFA': ada_expediente,
            'NCHI_HIDRAULICA': chi_hid,
            'NCHI_EXPLOTACION': chi_exp,
            'NCHI_VUELCO': chi_vue,
            'RED_MONITOREOS': red_monitoreos,
            'RENPRE_STATUS': renpre_status,
            'RESIDUOS_ESPECIALES_STATUS': rree_status,
            'CHE_STATUS': che_status,
            'AÑO_CHE': anio_che,
            'OBSERVACIONES_TICKETS_CONSULTA_CHE': obs_che,
            'GESTION_RESIDUOS_STATUS': gestion_res,
            'TIPO_RESIDUO': tipo_res,
            'ASP_STATUS': asp_status,
            'VENCIMIENTO_ASP': vto_asp,
            'EXPEDIENTE_ASP': expediente_asp,
            'VALVULAS_STATUS': valvulas_status,
            'VENCIMIENTO_CALIBRACION_ASP': vto_valvulas
        }

        try:
            template_bytes = template_file.read()
            out_bytes, filename = docgen.generate_report(template_bytes, user_data, st.session_state.hallazgos, st.session_state.muestreo_rows)
            st.success(f"Informe generado: {filename}")
            st.download_button("Descargar informe", data=out_bytes, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.exception(e)

st.markdown("---")
st.caption("APP generada a partir del código original en Colab / ipywidgets. Si necesitas que adapte textos, nombres de marcadores o la lógica de borrado condicional, dime y lo ajusto.")