"""
================================================================================
SYNCFIT — APLICACIÓN WEB DE ENTRENAMIENTO PERSONALIZADO
Versión completa con integración de base de datos Excel (Pandas) y Streamlit.
================================================================================
"""

import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ==============================================================================
# 1. CONFIGURACIÓN DE LA PÁGINA Y ESTILOS VISUALES (PREMIUM FITNESS TECH)
# ==============================================================================
st.set_page_config(
    page_title="SyncFit — Entrenamiento Personalizado",
    page_icon="💪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inyección de estilos CSS personalizados para lograr la paleta:
# Negro azulado (#101820), Azul oscuro (#172B4D), Azul eléctrico (#3978F6),
# Verde menta (#43D9B0), Gris claro (#F4F6FA), Gris medio (#737B8C)
st.markdown("""
<style>
    /* Estilos globales y tipografía */
    .main {
        background-color: #F4F6FA;
        font-family: 'Inter', sans-serif;
    }
    .stApp {
        background-color: #F4F6FA;
    }
    
    /* Contenedores y Tarjetas */
    .card-container {
        background-color: #FFFFFF;
        border-radius: 14px;
        padding: 24px;
        box-shadow: 0 4px 6px -1px rgba(16, 24, 32, 0.05), 0 2px 4px -1px rgba(16, 24, 32, 0.03);
        border: 1px solid #EAEFF7;
        margin-bottom: 20px;
    }
    
    /* Cabeceras y Títulos */
    h1, h2, h3 {
        color: #172B4D !important;
        font-weight: 700;
    }
    
    /* Botones principales */
    .stButton>button {
        background-color: #3978F6;
        color: #FFFFFF;
        border-radius: 10px;
        font-weight: 600;
        padding: 0.6rem 1.2rem;
        border: none;
        box-shadow: 0 2px 4px rgba(57, 120, 246, 0.2);
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #2860d8;
        box-shadow: 0 4px 8px rgba(57, 120, 246, 0.3);
    }
    
    /* Espacios reservados para recursos (Logotipo e Imágenes) */
    .placeholder-logo {
        border: 2px dashed #737B8C;
        background-color: #EAEFF7;
        color: #737B8C;
        text-align: center;
        padding: 10px;
        border-radius: 8px;
        font-size: 12px;
        font-weight: 600;
        margin-bottom: 10px;
    }
    .placeholder-image {
        border: 2px dashed #737B8C;
        background-color: #EAEFF7;
        color: #737B8C;
        display: flex;
        align-items: center;
        justify-content: center;
        height: 220px;
        border-radius: 12px;
        font-size: 14px;
        font-weight: 500;
        margin: 15px 0;
    }
    
    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #172B4D;
        color: #FFFFFF;
    }
    [data-testid="stSidebar"] .stMarkdown {
        color: #FFFFFF;
    }
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# 2. GESTIÓN DE ESTADO (SESSION STATE)
# ==============================================================================
if 'screen' not in st.session_state:
    st.session_state.screen = 'bienvenida'

if 'user_data' not in st.session_state:
    st.session_state.user_data = {
        'edad': 25,
        'peso': 70.0,
        'estatura': 175,
        'condicion': 'Intermedia',
        'experiencia': 'Intermedia',
        'objetivo': 'Perder peso',
        'tiempo': '30 min',
        'dias': '3 días',
        'lugar': 'En casa',
        'equipo': 'Ninguno',
        'aviso_aceptado': False
    }

if 'workout_history' not in st.session_state:
    st.session_state.workout_history = []

if 'active_workout' not in st.session_state:
    st.session_state.active_workout = {
        'in_progress': False,
        'current_index': 0,
        'routine_name': '',
        'exercises': []
    }

if 'reminders' not in st.session_state:
    st.session_state.reminders = {
        'active': True,
        'time': "08:00",
        'days': ["Lunes", "Miércoles", "Viernes"]
    }


# ==============================================================================
# 3. CARGA Y PROCESAMIENTO DE LA BASE DE DATOS EXCEL
# ==============================================================================
EXCEL_FILE = "excel luciano ejercicio 2 1.xlsx"

@st.cache_data
def load_excel_database():
    if not os.path.exists(EXCEL_FILE):
        # Fallback de seguridad por si el archivo no se encuentra en el directorio exacto
        data_fallback = {
            "Objetivo": ["Perder peso", "Perder peso", "Mantener fisicio", "Mantener fisicio", "Mejorar salud", "Mejorar salud", "Aumentar musculo", "Aumentar musculo"],
            "Ejercicios": ["Jumping Jacks", "Sentadillas", "Sentadillas", "Plancha", "Movilidad de cuello", "Caminata suave", "Sentadillas con mancuernas", "Flexiones de brazo"],
            "Ejecucion": [
                "Salta abriendo las piernas y levantando los brazos. Regresa a la posición inicial.",
                "Flexiona las rodillas llevando la cadera hacia atrás y vuelve a subir.",
                "Flexiona las rodillas llevando la cadera hacia atrás y vuelve a levantarte.",
                "Apoya los antebrazos y las puntas de los pies, manteniendo el cuerpo alineado.",
                "Gira lentamente la cabeza hacia ambos lados, sin forzar el movimiento.",
                "Camina a un ritmo cómodo, respirando con normalidad.",
                "Sostén las mancuernas a los lados, baja la cadera y vuelve a subir.",
                "Coloca las manos en el suelo, baja el pecho y empuja para regresar arriba."
            ]
        }
        return pd.DataFrame(data_fallback)
    
    try:
        df = pd.read_excel(EXCEL_FILE)
        # Rellenar valores vacíos (NaN) hacia abajo en la columna Objetivo
        if 'Objetivo' in df.columns:
            df['Objetivo'] = df['Objetivo'].ffill()
        return df
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
        return pd.DataFrame()

df_excel = load_excel_database()


# ==============================================================================
# 4. BARRA LATERAL DE NAVEGACIÓN (DASHBOARD MODERN / SIDEBAR)
# ==============================================================================
with st.sidebar:
    st.markdown('<div class="placeholder-logo">[ Logotipo SyncFit ]</div>', unsafe_allow_html=True)
    st.markdown("### **SyncFit**")
    st.markdown("---")
    
    # Menú de navegación lateral
    nav_option = st.radio(
        "Navegación Principal",
        [
            "🏠 Bienvenida",
            "📊 Panel Principal",
            "📋 Mis Rutinas",
            "⏱️ Entrenamiento en Curso",
            "📈 Mi Progreso",
            "⏰ Recordatorios",
            "👤 Mi Perfil",
            "🛡️ Avisos de Seguridad"
        ],
        index=0 if st.session_state.screen == 'bienvenida' else 1
    )
    
    # Sincronizar selección de radio con el session_state
    if "Bienvenida" in nav_option: st.session_state.screen = 'bienvenida'
    elif "Panel Principal" in nav_option: st.session_state.screen = 'dashboard'
    elif "Mis Rutinas" in nav_option: st.session_state.screen = 'rutinas'
    elif "Entrenamiento en Curso" in nav_option: st.session_state.screen = 'entrenamiento'
    elif "Mi Progreso" in nav_option: st.session_state.screen = 'progreso'
    elif "Recordatorios" in nav_option: st.session_state.screen = 'recordatorios'
    elif "Mi Perfil" in nav_option: st.session_state.screen = 'perfil'
    elif "Avisos de Seguridad" in nav_option: st.session_state.screen = 'seguridad'
    
    st.markdown("---")
    st.markdown("<small style='color: #737B8C;'>SyncFit v2.6 — 2026<br>Organización y constancia.</small>", unsafe_allow_html=True)


# ==============================================================================
# 5. IMPLEMENTACIÓN DE PANTALLAS Y FLUJO COMPLETO DE LA WEB APP
# ==============================================================================

# ------------------------------------------------------------------------------
# PANTALLA 1: BIENVENIDA
# ------------------------------------------------------------------------------
if st.session_state.screen == 'bienvenida':
    col1, col2 = st.columns([1.2, 1])
    
    with col1:
        st.markdown("# Tu entrenamiento. <br><span style='color: #3978F6;'>Tu tiempo. Tu ritmo.</span>", unsafe_allow_html=True)
        st.markdown("<p style='color: #737B8C; font-size: 1.1rem;'>Organiza tus entrenamientos con rutinas adaptadas a tus objetivos, tu experiencia y el tiempo que tienes disponible.</p>", unsafe_allow_html=True)
        
        st.markdown("""
        * Rutinas adaptadas a tus preferencias.
        * Entrenamientos organizados por tiempo.
        * Seguimiento de tus sesiones.
        """)
        
        if st.button("Crear mi rutina (Comenzar)", use_container_width=False):
            st.session_state.screen = 'datos_iniciales'
            st.rerun()
            
        st.markdown("<small style='color: #737B8C;'>Personaliza tu experiencia en unos sencillos pasos.</small>", unsafe_allow_html=True)
        
    with col2:
        st.markdown('<div class="placeholder-image">[ Espacio reservado para imagen principal de bienvenida ]</div>', unsafe_allow_html=True)
        
    st.markdown("---")
    
    # Tres tarjetas informativas inferiores
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown('<div class="card-container"><h4>⚙️ Personalización</h4><p style="font-size: 0.9rem; color: #737B8C;">Configura parámetros según tu experiencia y lugar de entreno.</p></div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="card-container"><h4>⏱️ Organización</h4><p style="font-size: 0.9rem; color: #737B8C;">Sesiones eficientes desde 15 hasta 60 minutos programadas.</p></div>', unsafe_allow_html=True)
    with c3:
        st.markdown('<div class="card-container"><h4>📈 Seguimiento</h4><p style="font-size: 0.9rem; color: #737B8C;">Registra la constancia de tus entrenamientos realizados.</p></div>', unsafe_allow_html=True)


# ------------------------------------------------------------------------------
# PANTALLA 2: DATOS INICIALES (PASO 1 DE 3)
# ------------------------------------------------------------------------------
elif st.session_state.screen == 'datos_iniciales':
    st.markdown("### Paso 1 de 3: Datos Iniciales")
    st.markdown("#### Vamos a personalizar tu experiencia")
    st.markdown("Cuéntanos un poco sobre ti para organizar mejor tus entrenamientos.")
    
    with st.form("form_datos"):
        edad = st.number_input("¿Qué edad tienes?", min_value=10, max_value=100, value=int(st.session_state.user_data['edad']))
        if edad < 18:
            st.warning("⚠️ Importante: Contar con orientación adecuada al realizar ejercicio siendo menor de edad.")
            
        col_p, col_e = st.columns(2)
        with col_p:
            peso = st.number_input("¿Cuál es tu peso actual (kg)?", min_value=30.0, max_value=200.0, value=float(st.session_state.user_data['peso']), step=0.5)
        with col_e:
            estatura = st.number_input("¿Cuánto mides (cm)?", min_value=100, max_value=220, value=int(st.session_state.user_data['estatura']))
            
        condicion = st.selectbox(
            "¿Cómo describirías tu condición física actual?",
            ["Principiante", "Intermedia", "Avanzada"],
            index=["Principiante", "Intermedia", "Avanzada"].index(st.session_state.user_data['condicion'])
        )
        
        experiencia = st.selectbox(
            "¿Qué tanta experiencia tienes con las rutinas de ejercicio?",
            ["Principiante", "Intermedia", "Avanzada"],
            index=["Principiante", "Intermedia", "Avanzada"].index(st.session_state.user_data['experiencia'])
        )
        
        submitted = st.form_submit_button("Continuar a Objetivos")
        if submitted:
            st.session_state.user_data['edad'] = edad
            st.session_state.user_data['peso'] = peso
            st.session_state.user_data['estatura'] = estatura
            st.session_state.user_data['condicion'] = condicion
            st.session_state.user_data['experiencia'] = experiencia
            st.session_state.screen = 'objetivos'
            st.rerun()


# ------------------------------------------------------------------------------
# PANTALLA 3: OBJETIVOS Y PERSONALIZACIÓN (PASO 2 DE 3)
# ------------------------------------------------------------------------------
elif st.session_state.screen == 'objetivos':
    st.markdown("### Paso 2 de 3: Objetivos y Preferencias")
    st.markdown("#### Diseñemos una rutina para ti")
    
    # Obtener objetivos disponibles del Excel
    objetivos_excel = df_excel['Objetivo'].unique().tolist() if not df_excel.empty else ["Perder peso", "Mantener fisicio", "Mejorar salud", "Aumentar musculo"]
    
    with st.form("form_objetivos"):
        objetivo = st.selectbox(
            "¿Cuál es tu principal objetivo?",
            objetivos_excel,
            index=0 if st.session_state.user_data['objetivo'] not in objetivos_excel else objetivos_excel.index(st.session_state.user_data['objetivo'])
        )
        
        col1, col2 = st.columns(2)
        with col1:
            tiempo = st.selectbox("¿Cuánto tiempo tienes para entrenar?", ["15 min", "30 min", "45 min", "60 min"], index=1)
            dias = st.selectbox("¿Cuántos días a la semana puedes entrenar?", ["2 días", "3 días", "4 días", "5 días"], index=1)
        with col2:
            lugar = st.selectbox("¿Dónde vas a entrenar principalmente?", ["En casa", "En gimnasio", "Al aire libre"], index=0)
            equipo = st.selectbox("¿Qué equipo tienes disponible?", ["Ninguno", "Mancuernas", "Bandas elásticas", "Gimnasio completo"], index=0)
            
        submitted_obj = st.form_submit_button("Revisar mi perfil")
        if submitted_obj:
            st.session_state.user_data['objetivo'] = objetivo
            st.session_state.user_data['tiempo'] = tiempo
            st.session_state.user_data['dias'] = dias
            st.session_state.user_data['lugar'] = lugar
            st.session_state.user_data['equipo'] = equipo
            st.session_state.screen = 'resumen'
            st.rerun()


# ------------------------------------------------------------------------------
# PANTALLA 4: RESUMEN Y CONFIRMACIÓN (PASO 3 DE 3)
# ------------------------------------------------------------------------------
elif st.session_state.screen == 'resumen':
    st.markdown("### Paso 3 de 3: Resumen del Perfil")
    st.markdown("#### Revisa tus preferencias")
    st.markdown("Comprueba que la información sea correcta antes de preparar tu rutina.")
    
    u = st.session_state.user_data
    
    st.markdown(f"""
    <div class="card-container">
        <ul style="line-height: 1.8; list-style-type: none; padding-left: 0;">
            <li><strong>Edad:</strong> {u['edad']} años</li>
            <li><strong>Peso y Estatura:</strong> {u['peso']} kg / {u['estatura']} cm</li>
            <li><strong>Condición física:</strong> {u['condicion']}</li>
            <li><strong>Experiencia:</strong> {u['experiencia']}</li>
            <li><strong>Objetivo principal:</strong> {u['objetivo']}</li>
            <li><strong>Tiempo disponible:</strong> {u['tiempo']}</li>
            <li><strong>Días por semana:</strong> {u['dias']}</li>
            <li><strong>Lugar de entrenamiento:</strong> {u['lugar']}</li>
            <li><strong>Equipo disponible:</strong> {u['equipo']}</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    col_back, col_next = st.columns(2)
    with col_back:
        if st.button("Editar datos"):
            st.session_state.screen = 'datos_iniciales'
            st.rerun()
    with col_next:
        if st.button("Continuar al Aviso de Seguridad"):
            st.session_state.screen = 'seguridad_previo'
            st.rerun()


# ------------------------------------------------------------------------------
# PANTALLA 5: AVISO DE SEGURIDAD PREVIO
# ------------------------------------------------------------------------------
elif st.session_state.screen == 'seguridad_previo':
    st.markdown("### Aviso antes de comenzar")
    
    st.markdown("""
    <div style="background-color: #FFF8E6; border: 1px solid #FFE0B2; padding: 16px; border-radius: 12px; font-size: 0.9rem; color: #5D4037; margin-bottom: 20px;">
    SyncFit es una herramienta de organización y orientación general del entrenamiento. Las rutinas se preparan a partir de las preferencias proporcionadas por el usuario y no sustituyen la valoración ni las recomendaciones individualizadas de un profesional de la salud. Los resultados pueden variar entre personas. SyncFit no garantiza cambios físicos, pérdida de peso, aumento de masa muscular ni mejoras específicas del rendimiento. Realiza los ejercicios de acuerdo con tus capacidades y utiliza una técnica adecuada. Si experimentas dolor inusual, mareo o malestar, detén la actividad.
    </div>
    """, unsafe_allow_html=True)
    
    aceptado = st.checkbox("He leído el aviso y entiendo el alcance de SyncFit.")
    
    if st.button("Generar mi rutina", disabled=not aceptado):
        st.session_state.user_data['aviso_aceptado'] = True
        st.success("¡Organizando tu entrenamiento correctamente desde tu base de datos!")
        st.session_state.screen = 'dashboard'
        st.rerun()


# ------------------------------------------------------------------------------
# PANTALLA 6: PANEL PRINCIPAL (DASHBOARD)
# ------------------------------------------------------------------------------
elif st.session_state.screen == 'dashboard':
    st.markdown("## Panel Principal")
    st.markdown(f"**¡Bienvenido a SyncFit!** Tu entrenamiento está organizado según tus preferencias para el objetivo: *{st.session_state.user_data['objetivo']}*.")
    
    # Tarjeta Principal del Entrenamiento de Hoy
    objetivo_actual = st.session_state.user_data['objetivo']
    ejercicios_rutina = df_excel[df_excel['Objetivo'] == objetivo_actual] if not df_excel.empty else pd.DataFrame()
    
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #172B4D 0%, #101820 100%); color: white; padding: 30px; border-radius: 16px; margin-bottom: 25px;">
        <span style="background-color: #43D9B0; color: #101820; padding: 4px 10px; border-radius: 20px; font-size: 11px; font-weight: bold;">RECOMENDADO HOY</span>
        <h2 style="color: white !important; margin-top: 10px;">Rutina: {objetivo_actual}</h2>
        <p style="color: #A6B0CF; font-size: 0.95rem;">Duración estimada: {st.session_state.user_data['tiempo']} | Nivel: {st.session_state.user_data['experiencia']} | Ejercicios totales: {len(ejercicios_rutina)}</p>
    </div>
    """, unsafe_allow_html=True)
    
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("Comenzar entrenamiento hoy"):
            st.session_state.active_workout['in_progress'] = True
            st.session_state.active_workout['current_index'] = 0
            st.session_state.active_workout['routine_name'] = objetivo_actual
            st.session_state.active_workout['exercises'] = ejercicios_rutina.to_dict('records')
            st.session_state.screen = 'entrenamiento'
            st.rerun()
    with col_btn2:
        if st.button("Ver rutina completa en detalle"):
            st.session_state.screen = 'rutinas'
            st.rerun()
            
    # Tarjetas secundarias de métricas
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f'<div class="card-container"><span style="color: #737B8C; font-size: 12px; font-weight: bold;">OBJETIVO</span><h3 style="margin: 5px 0 0 0; font-size: 1.2rem;">{objetivo_actual}</h3></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="card-container"><span style="color: #737B8C; font-size: 12px; font-weight: bold;">TIEMPO DISPONIBLE</span><h3 style="margin: 5px 0 0 0; font-size: 1.2rem;">{st.session_state.user_data["tiempo"]}</h3></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="card-container"><span style="color: #737B8C; font-size: 12px; font-weight: bold;">SESIONES COMPLETADAS</span><h3 style="margin: 5px 0 0 0; font-size: 1.2rem; color: #3978F6;">{len(st.session_state.workout_history)} entrenamientos</h3></div>', unsafe_allow_html=True)


# ------------------------------------------------------------------------------
# PANTALLA 7: MIS RUTINAS
# ------------------------------------------------------------------------------
elif st.session_state.screen == 'rutinas':
    st.markdown("## Mis Rutinas")
    st.markdown("Explora tus entrenamientos disponibles según la base de datos de ejercicios.")
    
    if not df_excel.empty:
        objetivos_disponibles = df_excel['Objetivo'].unique()
        obj_filtro = st.selectbox("Filtrar por Objetivo:", objetivos_disponibles)
        
        ej_filtrados = df_excel[df_excel['Objetivo'] == obj_filtro]
        
        st.markdown(f"### Detalle de la rutina: {obj_filtro}")
        st.markdown(f"**Duración estimada:** {st.session_state.user_data['tiempo']} | **Equipo:** {st.session_state.user_data['equipo']}")
        
        for idx, row in ej_filtrados.reset_index().iterrows():
            st.markdown(f"""
            <div class="card-container" style="padding: 16px;">
                <h4 style="margin: 0 0 8px 0; color: #172B4D;">{row['Ejercicios']}</h4>
                <p style="margin: 0; font-size: 0.9rem; color: #737B8C;"><strong>Ejecución:</strong> {row['Ejecucion']}</p>
                <div class="placeholder-image" style="height: 120px; margin-top: 10px;">[ Espacio reservado para imagen explicativa de {row['Ejercicios']} ]</div>
            </div>
            """, unsafe_allow_html=True)
            
        if st.button("Comenzar esta rutina"):
            st.session_state.active_workout['in_progress'] = True
            st.session_state.active_workout['current_index'] = 0
            st.session_state.active_workout['routine_name'] = obj_filtro
            st.session_state.active_workout['exercises'] = ej_filtrados.to_dict('records')
            st.session_state.screen = 'entrenamiento'
            st.rerun()
    else:
        st.info("No se encontraron registros en el archivo Excel.")


# ------------------------------------------------------------------------------
# PANTALLA 10: ENTRENAMIENTO EN CURSO
# ------------------------------------------------------------------------------
elif st.session_state.screen == 'entrenamiento':
    st.markdown("## Entrenamiento en Curso")
    
    workout = st.session_state.active_workout
    if not workout['in_progress'] or not workout['exercises']:
        st.warning("No hay ningún entrenamiento activo en este momento. Selecciona una rutina desde el panel o mis rutinas.")
        if st.button("Ir a Mis Rutinas"):
            st.session_state.screen = 'rutinas'
            st.rerun()
    else:
        idx = workout['current_index']
        exercises = workout['exercises']
        
        if idx < len(exercises):
            current_ex = exercises[idx]
            st.markdown(f"### Rutina: {workout['routine_name']}")
            st.progress((idx + 1) / len(exercises))
            st.markdown(f"**Ejercicio {idx + 1} de {len(exercises)}:** <span style='color: #3978F6;'>{current_ex['Ejercicios']}</span>", unsafe_allow_html=True)
            
            st.markdown(f"""
            <div class="card-container">
                <div class="placeholder-image">[ Espacio reservado para guía visual del ejercicio ]</div>
                <h4 style="color: #172B4D;">Instrucción de ejecución:</h4>
                <p style="font-size: 1rem; color: #3978F6;">{current_ex['Ejecucion']}</p>
            </div>
            """, unsafe_allow_html=True)
            
            col_prev, col_next = st.columns(2)
            with col_prev:
                if idx > 0 and st.button("← Ejercicio anterior"):
                    st.session_state.active_workout['current_index'] -= 1
                    st.rerun()
            with col_next:
                if idx < len(exercises) - 1:
                    if st.button("Marcar completado y siguiente →"):
                        st.session_state.active_workout['current_index'] += 1
                        st.rerun()
                else:
                    if st.button("Finalizar entrenamiento 🎉"):
                        st.session_state.workout_history.append({
                            'fecha': datetime.now().strftime("%Y-%m-%d %H:%M"),
                            'rutina': workout['routine_name'],
                            'ejercicios_total': len(exercises)
                        })
                        st.session_state.active_workout['in_progress'] = False
                        st.session_state.screen = 'entrenamiento_completado'
                        st.rerun()
        else:
            st.success("¡Has completado todos los ejercicios!")
            if st.button("Finalizar sesión"):
                st.session_state.screen = 'entrenamiento_completado'
                st.rerun()


# ------------------------------------------------------------------------------
# PANTALLA 11: ENTRENAMIENTO COMPLETADO
# ------------------------------------------------------------------------------
elif st.session_state.screen == 'entrenamiento_completado':
    st.markdown("## ¡Entrenamiento completado!")
    st.markdown("""
    <div class="card-container" style="text-align: center; padding: 40px;">
        <h1 style="color: #43D9B0 !important; font-size: 3rem;">✓</h1>
        <h3>¡Gran trabajo!</h3>
        <p style="color: #737B8C;">Has finalizado tu sesión correctamente. Tu constancia construye tus resultados.</p>
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("Volver al Panel Principal"):
        st.session_state.screen = 'dashboard'
        st.rerun()


# ------------------------------------------------------------------------------
# PANTALLA 12: MI PROGRESO
# ------------------------------------------------------------------------------
elif st.session_state.screen == 'progreso':
    st.markdown("## Mi Progreso")
    st.markdown("Consulta tus entrenamientos completados y observa tu constancia.")
    
    total_sesiones = len(st.session_state.workout_history)
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"""
        <div class="card-container">
            <span style="color: #737B8C; font-size: 12px; font-weight: bold;">ENTRENAMIENTOS COMPLETADOS</span>
            <h2 style="color: #3978F6; margin: 10px 0 0 0;">{total_sesiones}</h2>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="card-container">
            <span style="color: #737B8C; font-size: 12px; font-weight: bold;">ESTADO DE CONSTANCIA</span>
            <h2 style="color: #43D9B0; margin: 10px 0 0 0;">{"Activo" if total_sesiones > 0 else "Iniciado"}</h2>
        </div>
        """, unsafe_allow_html=True)
        
    if total_sesiones > 0:
        st.markdown("### Historial de Sesiones")
        df_history = pd.DataFrame(st.session_state.workout_history)
        st.dataframe(df_history, use_container_width=True, hide_index=True)
    else:
        st.info("Tu progreso aparecerá aquí cuando completes tus primeros entrenamientos.")


# ------------------------------------------------------------------------------
# PANTALLA 13: RECORDATORIOS
# ------------------------------------------------------------------------------
elif st.session_state.screen == 'recordatorios':
    st.markdown("## Recordatorios")
    st.markdown("Organiza tus preferencias para mantener tus sesiones programadas.")
    
    with st.form("form_recordatorios"):
        activo = st.checkbox("Activar recordatorios visuales", value=st.session_state.reminders['active'])
        hora = st.text_input("Hora del entrenamiento", value=st.session_state.reminders['time'])
        dias_sel = st.multiselect("Días de entrenamiento", ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"], default=st.session_state.reminders['days'])
        
        guardar = st.form_submit_button("Guardar preferencias")
        if guardar:
            st.session_state.reminders['active'] = activo
            st.session_state.reminders['time'] = hora
            st.session_state.reminders['days'] = dias_sel
            st.success("¡Preferencias de recordatorios guardadas correctamente!")


# ------------------------------------------------------------------------------
# PANTALLA 14: MI PERFIL Y CONFIGURACIÓN
# ------------------------------------------------------------------------------
elif st.session_state.screen == 'perfil':
    st.markdown("## Mi Perfil y Configuración")
    u = st.session_state.user_data
    
    st.markdown(f"""
    <div class="card-container">
        <h3>Datos Personales</h3>
        <p><strong>Edad:</strong> {u['edad']} años</p>
        <p><strong>Peso:</strong> {u['peso']} kg</p>
        <p><strong>Estatura:</strong> {u['estatura']} cm</p>
    </div>
    
    <div class="card-container">
        <h3>Preferencias de Entrenamiento</h3>
        <p><strong>Objetivo:</strong> {u['objetivo']}</p>
        <p><strong>Experiencia:</strong> {u['experiencia']}</p>
        <p><strong>Condición:</strong> {u['condicion']}</p>
        <p><strong>Tiempo disponible:</strong> {u['tiempo']}</p>
        <p><strong>Lugar:</strong> {u['lugar']} ({u['equipo']})</p>
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("Modificar mis preferencias"):
        st.session_state.screen = 'datos_iniciales'
        st.rerun()


# ------------------------------------------------------------------------------
# PANTALLA 21: AVISOS DE SEGURIDAD Y ALCANCE
# ------------------------------------------------------------------------------
elif st.session_state.screen == 'seguridad':
    st.markdown("## Entrena con información clara")
    st.markdown("""
    <div class="card-container">
        <p style="line-height: 1.7; color: #172B4D;">
        SyncFit es una herramienta de organización y orientación general del entrenamiento. No sustituye la valoración, el diagnóstico ni las recomendaciones individualizadas de un médico, fisioterapeuta u otro profesional de la salud. Los resultados pueden variar entre personas. SyncFit no garantiza cambios físicos, pérdida de peso, aumento de masa muscular ni mejoras específicas del rendimiento. Realiza los ejercicios de acuerdo con tus capacidades. Si experimentas dolor inusual, mareo o malestar, detén la actividad. Si tienes una lesión, una condición médica o dudas sobre qué actividad puedes realizar, solicita orientación profesional.
        </p>
        <h4 style="margin-top: 20px;">Uso de la información</h4>
        <p style="color: #737B8C;">La información proporcionada se utiliza únicamente para configurar la experiencia de entrenamiento en tu sesión actual.</p>
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("Volver al Panel"):
        st.session_state.screen = 'dashboard'
        st.rerun()