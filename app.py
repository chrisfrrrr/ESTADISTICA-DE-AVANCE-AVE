from __future__ import annotations

import os
import re
from datetime import datetime, date, timezone
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
import plotly.express as px
from dotenv import load_dotenv

from core.canvas_client import CanvasClient, CanvasAPIError
from core.analytics import (
    normalize_enrollments,
    normalize_assignments,
    normalize_submissions,
    build_student_summary,
    normalize_modules,
    normalize_module_items,
    build_module_completion_matrix,
)
from core.reports import excel_bytes, pdf_bytes, individual_pdf_bytes, DEV

load_dotenv()
st.set_page_config(page_title='AVE Monitor Académico Pro 4.0 FIX2', page_icon='assets/app_icon.ico', layout='wide')

LOGO_AVE = 'assets/logo_ave.png'
LOGO_UVG = 'assets/logo_uvg.png'
RISK_ORDER = ['Bajo', 'Medio', 'Alto']

st.markdown('''
<style>
.block-container {padding-top: 1.0rem; max-width: 1550px;}
[data-testid="stSidebar"] {background: linear-gradient(180deg, #F3F7FC 0%, #EAF0F7 100%);} 
.ave-title {font-size: 2.2rem; font-weight: 900; color:#172B85; margin-bottom:0; letter-spacing:-.02em;}
.ave-subtitle {font-size: 1rem; color:#475569; margin-top:.2rem; margin-bottom:1.0rem;}
.ave-card {border:1px solid #E5E7EB;border-radius:20px;padding:18px;background:#fff;box-shadow:0 8px 28px rgba(15,23,42,.07);}
.kpi-card {border:1px solid #E5E7EB;border-radius:18px;padding:16px;background:#fff;box-shadow:0 2px 12px rgba(15,23,42,.06)}
.kpi-label {font-size:.82rem;color:#64748B;}
.kpi-value {font-size:1.65rem;font-weight:800;color:#172B85;}
.section-note {background:#F8FAFC;border-left:5px solid #00A83B;padding:12px;border-radius:12px;color:#334155;}
.footer {font-size:.80rem;color:#64748B;text-align:center;margin-top:20px;}
</style>
''', unsafe_allow_html=True)

for key, default in {'client': None, 'courses': [], 'analysis': None, 'last_user': None}.items():
    if key not in st.session_state:
        st.session_state[key] = default


def clean_course_group(name: str) -> str:
    """Agrupa nombres tipo 'Matemáticas - SECCIÓN - 10 - 2026 - 1' como 'Matemáticas'."""
    s = str(name or 'Curso sin nombre').strip()
    patterns = [
        r'\s*[-–—]\s*SECCI[ÓO]N\s*[-–—]?\s*\d+.*$',
        r'\s*[-–—]\s*SECTION\s*[-–—]?\s*\d+.*$',
        r'\s*[-–—]\s*SEC\.?\s*\d+.*$',
    ]
    for p in patterns:
        s2 = re.sub(p, '', s, flags=re.IGNORECASE).strip()
        if s2 != s:
            return s2
    # fallback: si termina con varios bloques numéricos de periodo, conserva el nombre base
    s2 = re.sub(r'\s*[-–—]\s*\d+\s*[-–—]\s*20\d{2}\s*[-–—]\s*\d+.*$', '', s).strip()
    return s2 or s


def safe_int_id(value):
    try:
        if pd.isna(value):
            return None
        return int(value)
    except Exception:
        return None


def diagnostic_rows_to_df(rows: list[dict], active_user_ids: set, course_id, course_name, course_group) -> pd.DataFrame:
    out = []
    for e in rows or []:
        u = e.get('user') or {}
        uid = u.get('id') or e.get('user_id')
        uid_int = safe_int_id(uid)
        correo = u.get('login_id') or u.get('email') or ''
        estado = e.get('enrollment_state') or e.get('workflow_state') or e.get('_diagnostic_state_requested') or 'sin_estado'
        if uid_int in active_user_ids:
            clasif = 'Activo analizado'
        elif str(estado).lower() in {'invited', 'creation_pending', 'pending'}:
            clasif = 'Pendiente de registro'
        elif str(estado).lower() in {'inactive', 'deleted', 'rejected'}:
            clasif = 'No activo / retirado'
        elif str(estado).lower() in {'completed'}:
            clasif = 'Curso concluido'
        else:
            clasif = 'Detectado no analizado'
        if (not correo) or str(correo).strip().lower() in {'pendiente', 'pending', 'none', 'nan'}:
            clasif = 'Datos incompletos / no registrado'
        out.append({
            'course_id': course_id,
            'curso_canvas': course_name,
            'curso_general': course_group,
            'user_id': uid,
            'estudiante': u.get('sortable_name') or u.get('name') or 'Sin nombre',
            'nombre': u.get('name') or u.get('sortable_name') or 'Sin nombre',
            'correo': correo,
            'estado_canvas': estado,
            'section_id': e.get('course_section_id'),
            'tipo_inscripcion': e.get('type'),
            'clasificacion_registro': clasif,
            'ultima_actividad_canvas': e.get('last_activity_at'),
        })
    return pd.DataFrame(out)


def append_non_active_to_summary(summary: pd.DataFrame, diag_df: pd.DataFrame, analysis_dt: datetime) -> pd.DataFrame:
    """Incluye estudiantes detectados por Canvas pero no activos/analizados, con clasificación separada."""
    if diag_df is None or diag_df.empty:
        if 'clasificacion_registro' not in summary.columns:
            summary['clasificacion_registro'] = 'Activo analizado'
        return summary
    base = summary.copy()
    if 'clasificacion_registro' not in base.columns:
        base['clasificacion_registro'] = 'Activo analizado'
    active_ids = set(base['user_id'].dropna().astype(str)) if not base.empty and 'user_id' in base.columns else set()
    missing = diag_df[~diag_df['user_id'].astype(str).isin(active_ids)].copy()
    if missing.empty:
        return base
    extra = pd.DataFrame({
        'user_id': missing['user_id'],
        'estudiante': missing['estudiante'],
        'nombre': missing['nombre'],
        'correo': missing['correo'],
        'section_id': missing.get('section_id', ''),
        'ultima_actividad': pd.NaT,
        'horas_sin_actividad': None,
        'riesgo_desconexion': 'No aplica',
        'tiempo_total_horas': 0.0,
        'horas_esperadas': base['horas_esperadas'].max() if 'horas_esperadas' in base.columns and not base.empty else 0,
        'deficit_horas': 0.0,
        'cumplimiento_horas': 'No aplica',
        'actividades_total': 0,
        'actividades_actuales': 0,
        'entregadas': 0,
        'entregadas_actuales': 0,
        'pendientes': 0,
        'pendientes_actuales': 0,
        'pendientes_futuros': 0,
        'atrasadas': 0,
        'porcentaje_avance': 0.0,
        'porcentaje_avance_curso': 0.0,
        'promedio_score': 0.0,
        'puntaje_riesgo': 0,
        'riesgo_integral': 'No aplica',
        'segmento_ave': missing['clasificacion_registro'],
        'accion_recomendada': 'Validar registro o estado de inscripción en Canvas',
        'clasificacion_registro': missing['clasificacion_registro'],
        'estado_canvas': missing['estado_canvas'],
    })
    for col in base.columns:
        if col not in extra.columns:
            extra[col] = ''
    for col in extra.columns:
        if col not in base.columns:
            base[col] = ''
    return pd.concat([base, extra[base.columns]], ignore_index=True)


def adjust_completion_risk(df: pd.DataFrame, completion_threshold: float, current_threshold: float) -> pd.DataFrame:
    out = df.copy()
    if out.empty:
        return out
    for col in ['porcentaje_avance', 'porcentaje_avance_curso', 'pendientes_actuales', 'atrasadas', 'riesgo_integral']:
        if col not in out.columns:
            out[col] = 0 if col != 'riesgo_integral' else 'Bajo'
    final_mask = (
        (pd.to_numeric(out['porcentaje_avance_curso'], errors='coerce').fillna(0) >= completion_threshold)
        | ((pd.to_numeric(out['porcentaje_avance'], errors='coerce').fillna(0) >= current_threshold)
           & (pd.to_numeric(out['pendientes_actuales'], errors='coerce').fillna(0) == 0)
           & (pd.to_numeric(out['atrasadas'], errors='coerce').fillna(0) == 0))
    ) & out.get('clasificacion_registro', 'Activo analizado').eq('Activo analizado')
    out['estado_finalizacion'] = 'En desarrollo'
    out.loc[final_mask, 'estado_finalizacion'] = 'Finalizado o al día por entregas'
    out.loc[final_mask, 'riesgo_integral'] = 'Bajo'
    out.loc[final_mask, 'segmento_ave'] = 'Finalizado / al día'
    out.loc[final_mask, 'accion_recomendada'] = 'Monitoreo regular; no priorizar por desconexión si mantiene entregables completos'
    out.loc[final_mask, 'puntaje_riesgo'] = out.loc[final_mask, 'puntaje_riesgo'].apply(lambda x: min(int(x or 0), 20))
    out['alerta_72h'] = pd.to_numeric(out.get('horas_sin_actividad', pd.Series([0]*len(out))), errors='coerce').fillna(0).ge(72).map({True: 'Sí', False: 'No'})
    out.loc[final_mask, 'alerta_72h'] = 'No prioritaria por entregables'
    return out


def process_one_course(client: CanvasClient, course: dict, analysis_dt: datetime, course_start: date, daily_hours: float, only_business: bool, completion_threshold: float, current_threshold: float) -> dict:
    course_id = course.get('id')
    course_name = course.get('name', 'Curso')
    course_group = clean_course_group(course_name)

    enrollments = client.enrollments(course_id, None)
    enroll_df = normalize_enrollments(enrollments, analysis_dt, course_start, daily_hours, only_business)
    valid_ids = set(enroll_df['user_id'].dropna().astype(int).tolist()) if not enroll_df.empty else set()

    try:
        diag_raw = client.enrollments_for_diagnostic(course_id, None)
    except Exception:
        diag_raw = enrollments
    diag_df = diagnostic_rows_to_df(diag_raw, valid_ids, course_id, course_name, course_group)

    assignments = client.assignments(course_id)
    assign_df = normalize_assignments(assignments)
    modules_raw = client.modules(course_id)
    modules_df = normalize_modules(modules_raw)
    module_items_df = normalize_module_items(modules_raw)
    submissions = client.submissions(course_id, student_ids=sorted(valid_ids) if valid_ids else None, chunk_size=10)
    sub_df = normalize_submissions(submissions)
    if valid_ids and not sub_df.empty:
        sub_df = sub_df[sub_df['user_id'].isin(valid_ids)]
    module_matrix = build_module_completion_matrix(enroll_df, module_items_df, sub_df, analysis_dt)
    summary = build_student_summary(enroll_df, sub_df, analysis_dt)
    summary = append_non_active_to_summary(summary, diag_df, analysis_dt)
    summary = adjust_completion_risk(summary, completion_threshold, current_threshold)

    # Agregar metadatos del curso sin duplicar columnas.
    # Algunas tablas, especialmente diagnostic y summary ampliado con estudiantes no activos,
    # ya pueden traer curso_general/curso_canvas/course_id. Usar insert() directamente
    # provoca el error: "cannot insert curso_general, already exists".
    def add_course_meta(df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return df

        meta = {
            'curso_general': course_group,
            'curso_canvas': course_name,
            'course_id': course_id,
        }

        for col, value in meta.items():
            if col not in df.columns:
                df[col] = value
            else:
                df[col] = df[col].replace('', pd.NA).fillna(value)

        first_cols = ['curso_general', 'curso_canvas', 'course_id']
        other_cols = [c for c in df.columns if c not in first_cols]
        return df[first_cols + other_cols]

    summary = add_course_meta(summary)
    sub_df = add_course_meta(sub_df)
    assign_df = add_course_meta(assign_df)
    modules_df = add_course_meta(modules_df)
    module_items_df = add_course_meta(module_items_df)
    module_matrix = add_course_meta(module_matrix)
    diag_df = add_course_meta(diag_df)

    return {
        'course_id': course_id,
        'course_name': course_name,
        'course_group': course_group,
        'summary': summary,
        'submissions': sub_df,
        'assignments': assign_df,
        'modules': modules_df,
        'module_items': module_items_df,
        'module_matrix': module_matrix,
        'diagnostic': diag_df,
    }


def concat_frames(results: list[dict], key: str) -> pd.DataFrame:
    frames = [r.get(key) for r in results if isinstance(r.get(key), pd.DataFrame) and not r.get(key).empty]
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def course_global_summary(summary: pd.DataFrame) -> pd.DataFrame:
    if summary.empty:
        return pd.DataFrame()
    df = summary.copy()
    for c in ['riesgo_integral', 'clasificacion_registro', 'estado_finalizacion']:
        if c not in df.columns:
            df[c] = ''
    def count_eq(col, value):
        return (col == value).sum()
    rows = []
    for cg, g in df.groupby('curso_general', dropna=False):
        activos = g[g['clasificacion_registro'].eq('Activo analizado')]
        rows.append({
            'Curso general': cg,
            'Secciones/cursos Canvas': g['curso_canvas'].nunique(),
            'Estudiantes detectados': g['user_id'].nunique(),
            'Activos analizados': len(activos),
            'No registrados/datos incompletos': int(g['clasificacion_registro'].astype(str).str.contains('no registrado|incompletos', case=False, na=False).sum()),
            'Pendientes de registro': int(g['clasificacion_registro'].eq('Pendiente de registro').sum()),
            'Finalizados o al día': int(g['estado_finalizacion'].eq('Finalizado o al día por entregas').sum()),
            'Riesgo alto': int(g['riesgo_integral'].eq('Alto').sum()),
            'Riesgo medio': int(g['riesgo_integral'].eq('Medio').sum()),
            'Riesgo bajo': int(g['riesgo_integral'].eq('Bajo').sum()),
            'Avance actual promedio %': round(float(pd.to_numeric(activos.get('porcentaje_avance', pd.Series(dtype=float)), errors='coerce').mean() or 0), 1) if not activos.empty else 0,
            'Avance curso promedio %': round(float(pd.to_numeric(activos.get('porcentaje_avance_curso', pd.Series(dtype=float)), errors='coerce').mean() or 0), 1) if not activos.empty else 0,
            'Pendientes actuales': int(pd.to_numeric(activos.get('pendientes_actuales', pd.Series(dtype=float)), errors='coerce').fillna(0).sum()) if not activos.empty else 0,
            'Atrasadas': int(pd.to_numeric(activos.get('atrasadas', pd.Series(dtype=float)), errors='coerce').fillna(0).sum()) if not activos.empty else 0,
        })
    return pd.DataFrame(rows).sort_values(['Riesgo alto', 'Pendientes actuales'], ascending=[False, False])


with st.sidebar:
    cols = st.columns(2)
    if Path(LOGO_AVE).exists():
        cols[0].image(LOGO_AVE, use_container_width=True)
    if Path(LOGO_UVG).exists():
        cols[1].image(LOGO_UVG, use_container_width=True)
    st.markdown('### Configuración Canvas')
    st.caption('AVE Monitor Académico Pro 4.0')
    canvas_url = st.text_input('URL Canvas', value=os.getenv('CANVAS_URL', 'https://uvg.instructure.com'))
    token = st.text_input('Token de acceso', value=os.getenv('CANVAS_TOKEN', ''), type='password')
    generated_by = st.text_input('Nombre de quien genera el informe', value='')
    st.divider()
    st.markdown('### Parámetros académicos')
    daily_hours = st.number_input('Meta mínima diaria de conexión (horas)', min_value=0.5, max_value=12.0, value=2.0, step=0.5)
    course_start = st.date_input('Fecha de inicio del curso', value=date.today())
    only_business = st.checkbox('Calcular meta solo con días hábiles', value=False)
    analysis_date = st.date_input('Fecha de corte del análisis', value=date.today())
    completion_threshold = st.slider('Porcentaje para considerar curso finalizado', min_value=70, max_value=100, value=100, step=5)
    current_threshold = st.slider('Porcentaje para considerar estudiante al día', min_value=60, max_value=100, value=90, step=5)
    st.divider()
    if st.button('Probar conexión / cargar cursos', use_container_width=True, type='primary'):
        try:
            c = CanvasClient(canvas_url, token)
            me = c.whoami()
            courses = c.courses()
            st.session_state.client = c
            st.session_state.courses = courses
            st.session_state.last_user = me.get('name', 'usuario Canvas')
            st.success(f'Conexión correcta: {st.session_state.last_user}')
        except Exception as e:
            st.error(f'No se pudo conectar: {e}')
    st.caption(f'Desarrollador: {DEV}')

st.markdown('<p class="ave-title">AVE Monitor Académico Pro 4.0</p>', unsafe_allow_html=True)
st.markdown('<p class="ave-subtitle">Análisis global por Canvas: conexión, entregables, estudiantes no registrados, riesgo ajustado, Excel y PDF ejecutivo institucional.</p>', unsafe_allow_html=True)

if not st.session_state.client:
    st.info('Ingrese URL de Canvas y token en la barra lateral. Luego presione “Probar conexión / cargar cursos”.')
    st.stop()

client: CanvasClient = st.session_state.client
courses = st.session_state.courses or []
if not courses:
    st.warning('No se encontraron cursos activos con este token.')
    st.stop()

course_options = {f"{c.get('name','Sin nombre')} | ID {c.get('id')}": c for c in courses}
labels = list(course_options.keys())
with st.container():
    st.markdown('<div class="ave-card">', unsafe_allow_html=True)
    st.markdown('### Selección de cursos y secciones Canvas')
    st.caption('Seleccione todas las secciones/cursos Canvas que pertenecen al periodo AVE. La app agrupa automáticamente por nombre general del curso.')
    col1, col2 = st.columns([3, 1])
    with col1:
        selected_labels = st.multiselect('Cursos Canvas a incluir en el informe global', labels, default=[])
    with col2:
        st.write('')
        st.write('')
        if st.button('Seleccionar todos', use_container_width=True):
            selected_labels = labels
            st.session_state['_selected_all_info'] = True
            st.info('Seleccione manualmente los cursos si Streamlit no actualiza la lista en pantalla.')
    st.markdown('</div>', unsafe_allow_html=True)

selected_courses = [course_options[l] for l in selected_labels]
if selected_courses:
    preview = pd.DataFrame([{'Curso Canvas': c.get('name'), 'ID': c.get('id'), 'Curso general detectado': clean_course_group(c.get('name'))} for c in selected_courses])
    with st.expander('Vista previa de agrupación automática', expanded=True):
        st.dataframe(preview, use_container_width=True, hide_index=True)

cA, cB = st.columns([1.2, 4])
generate = cA.button('Generar análisis global', type='primary', use_container_width=True)
if cB.button('Limpiar resultados', use_container_width=False):
    st.session_state.analysis = None
    st.rerun()

if generate:
    if not selected_courses:
        st.error('Seleccione al menos un curso o sección Canvas para analizar.')
    else:
        progress = st.progress(0, text='Iniciando consulta a Canvas...')
        results = []
        errors = []
        analysis_dt = datetime.combine(analysis_date, datetime.max.time()).replace(tzinfo=timezone.utc)
        total_courses = len(selected_courses)
        for idx, course in enumerate(selected_courses, start=1):
            name = course.get('name', 'Curso')
            pct = int((idx - 1) / max(total_courses, 1) * 90)
            progress.progress(pct, text=f'Procesando {idx}/{total_courses}: {name}')
            try:
                results.append(process_one_course(client, course, analysis_dt, course_start, daily_hours, only_business, float(completion_threshold), float(current_threshold)))
            except Exception as e:
                errors.append({'curso': name, 'id': course.get('id'), 'error': str(e)})
        summary = concat_frames(results, 'summary')
        submissions = concat_frames(results, 'submissions')
        assignments = concat_frames(results, 'assignments')
        modules = concat_frames(results, 'modules')
        module_items = concat_frames(results, 'module_items')
        module_matrix = concat_frames(results, 'module_matrix')
        diagnostic = concat_frames(results, 'diagnostic')
        global_course_df = course_global_summary(summary)
        st.session_state.analysis = {
            'results': results,
            'summary': summary,
            'submissions': submissions,
            'assignments': assignments,
            'modules': modules,
            'module_items': module_items,
            'module_matrix': module_matrix,
            'diagnostic': diagnostic,
            'global_course_df': global_course_df,
            'errors': errors,
            'analysis_date': str(analysis_date),
            'generated_by': generated_by,
            'selected_count': len(selected_courses),
        }
        progress.progress(100, text='Análisis global finalizado')
        if errors:
            st.warning(f'Análisis generado con {len(errors)} curso(s) con error. Revise la pestaña Diagnóstico.')
        else:
            st.success('Análisis global generado correctamente.')

analysis = st.session_state.analysis
if not analysis:
    st.stop()

summary = analysis['summary']
sub_df = analysis['submissions']
assign_df = analysis['assignments']
modules_df = analysis['modules']
module_items_df = analysis['module_items']
module_matrix = analysis['module_matrix']
diagnostic_df = analysis['diagnostic']
global_course_df = analysis['global_course_df']
errors = analysis.get('errors', [])

if summary.empty:
    st.warning('No se obtuvo información de estudiantes con los cursos seleccionados.')
    if errors:
        st.dataframe(pd.DataFrame(errors), use_container_width=True, hide_index=True)
    st.stop()

# KPIs generales
active = summary[summary.get('clasificacion_registro', pd.Series(['Activo analizado']*len(summary))).eq('Activo analizado')].copy()
counts = summary['riesgo_integral'].value_counts().to_dict() if 'riesgo_integral' in summary.columns else {}
finalizados = int(summary.get('estado_finalizacion', pd.Series(dtype=str)).eq('Finalizado o al día por entregas').sum()) if 'estado_finalizacion' in summary.columns else 0
no_reg = int(summary.get('clasificacion_registro', pd.Series(dtype=str)).astype(str).str.contains('no registrado|incompletos', case=False, na=False).sum()) if 'clasificacion_registro' in summary.columns else 0
pend_reg = int(summary.get('clasificacion_registro', pd.Series(dtype=str)).eq('Pendiente de registro').sum()) if 'clasificacion_registro' in summary.columns else 0

k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
k1.metric('Cursos/secciones Canvas', analysis.get('selected_count', 0))
k2.metric('Cursos generales', summary['curso_general'].nunique() if 'curso_general' in summary else 0)
k3.metric('Estudiantes detectados', summary['user_id'].nunique() if 'user_id' in summary else len(summary))
k4.metric('Activos analizados', len(active))
k5.metric('Riesgo alto', counts.get('Alto', 0))
k6.metric('Finalizados/al día', finalizados)
k7.metric('No registrados', no_reg + pend_reg)

st.markdown('<div class="section-note"><b>Lectura operativa:</b> el riesgo ya se ajusta con entregables. Un estudiante con baja conexión puede dejar de priorizarse si ya terminó o está al día con las actividades.</div>', unsafe_allow_html=True)

tab_res, tab_cursos, tab_riesgo, tab_est, tab_ent, tab_diag, tab_rep = st.tabs([
    'Resumen ejecutivo', 'Cursos globales', 'Riesgo y alertas', 'Estudiantes', 'Entregables', 'Diagnóstico', 'Exportables'
])

with tab_res:
    c1, c2 = st.columns(2)
    with c1:
        risk_plot = summary[summary['riesgo_integral'].isin(['Bajo','Medio','Alto'])]
        if not risk_plot.empty:
            fig = px.pie(risk_plot, names='riesgo_integral', title='Distribución del riesgo integral ajustado', hole=.42)
            st.plotly_chart(fig, use_container_width=True)
    with c2:
        if 'estado_finalizacion' in summary.columns:
            fig = px.histogram(summary, x='estado_finalizacion', title='Estado por entregables', text_auto=True)
            st.plotly_chart(fig, use_container_width=True)
    c3, c4 = st.columns(2)
    with c3:
        if not global_course_df.empty:
            fig = px.bar(global_course_df, x='Curso general', y='Riesgo alto', title='Riesgo alto por curso general', text='Riesgo alto')
            st.plotly_chart(fig, use_container_width=True)
    with c4:
        if not global_course_df.empty:
            fig = px.bar(global_course_df, x='Curso general', y='Finalizados o al día', title='Estudiantes finalizados o al día por curso', text='Finalizados o al día')
            st.plotly_chart(fig, use_container_width=True)

with tab_cursos:
    st.markdown('### Informe global por curso')
    st.dataframe(global_course_df, use_container_width=True, hide_index=True)
    st.markdown('### Distribución por curso Canvas/sección')
    view_cols = [c for c in ['curso_general','curso_canvas','estudiante','correo','clasificacion_registro','estado_finalizacion','riesgo_integral','porcentaje_avance','porcentaje_avance_curso','pendientes_actuales','atrasadas','horas_sin_actividad','accion_recomendada'] if c in summary.columns]
    st.dataframe(summary[view_cols], use_container_width=True, hide_index=True)

with tab_riesgo:
    st.markdown('### Priorización de estudiantes')
    mode = st.radio('Vista', ['Todos', 'Solo alto y medio', 'Intervención real', 'Finalizados/al día', 'No registrados o pendientes'], horizontal=True)
    rv = summary.copy()
    if mode == 'Solo alto y medio':
        rv = rv[rv['riesgo_integral'].isin(['Alto','Medio'])]
    elif mode == 'Intervención real':
        rv = rv[(rv['riesgo_integral'].isin(['Alto','Medio'])) & (~rv.get('estado_finalizacion', pd.Series(['']*len(rv))).eq('Finalizado o al día por entregas'))]
    elif mode == 'Finalizados/al día':
        rv = rv[rv.get('estado_finalizacion', pd.Series(['']*len(rv))).eq('Finalizado o al día por entregas')]
    elif mode == 'No registrados o pendientes':
        rv = rv[~rv.get('clasificacion_registro', pd.Series(['Activo analizado']*len(rv))).eq('Activo analizado')]
    if 'puntaje_riesgo' in rv.columns:
        rv = rv.sort_values(['puntaje_riesgo'], ascending=False)
    cols = [c for c in ['curso_general','curso_canvas','estudiante','correo','clasificacion_registro','estado_finalizacion','riesgo_integral','segmento_ave','puntaje_riesgo','horas_sin_actividad','alerta_72h','pendientes_actuales','pendientes_futuros','atrasadas','porcentaje_avance','porcentaje_avance_curso','accion_recomendada'] if c in rv.columns]
    st.dataframe(rv[cols], use_container_width=True, hide_index=True)
    st.download_button('Descargar alertas CSV', rv[cols].to_csv(index=False).encode('utf-8-sig'), file_name='alertas_globales_ave.csv', mime='text/csv', use_container_width=True)

with tab_est:
    st.markdown('### Base completa consolidada')
    f1, f2, f3 = st.columns(3)
    cursos_filter = f1.multiselect('Curso general', sorted(summary['curso_general'].dropna().unique().tolist()), default=sorted(summary['curso_general'].dropna().unique().tolist())) if 'curso_general' in summary else []
    riesgos_filter = f2.multiselect('Riesgo', sorted(summary['riesgo_integral'].dropna().unique().tolist()), default=sorted(summary['riesgo_integral'].dropna().unique().tolist())) if 'riesgo_integral' in summary else []
    reg_filter = f3.multiselect('Clasificación registro', sorted(summary['clasificacion_registro'].dropna().unique().tolist()), default=sorted(summary['clasificacion_registro'].dropna().unique().tolist())) if 'clasificacion_registro' in summary else []
    fv = summary.copy()
    if cursos_filter:
        fv = fv[fv['curso_general'].isin(cursos_filter)]
    if riesgos_filter:
        fv = fv[fv['riesgo_integral'].isin(riesgos_filter)]
    if reg_filter:
        fv = fv[fv['clasificacion_registro'].isin(reg_filter)]
    st.dataframe(fv, use_container_width=True, hide_index=True)

with tab_ent:
    st.markdown('### Entregables y submissions')
    c1, c2 = st.columns(2)
    with c1:
        st.caption('Actividades publicadas detectadas')
        st.dataframe(assign_df, use_container_width=True, hide_index=True)
    with c2:
        st.caption('Entregas recibidas desde Canvas')
        st.dataframe(sub_df, use_container_width=True, hide_index=True)
    if not sub_df.empty and 'curso_general' in sub_df.columns:
        ent = sub_df.copy()
        ent['entregada'] = ent['submitted_at'].notna() | ent['workflow_state'].isin(['submitted','graded'])
        ent_sum = ent.groupby('curso_general').agg(Registros=('assignment_id','count'), Entregadas=('entregada','sum')).reset_index()
        fig = px.bar(ent_sum, x='curso_general', y=['Registros','Entregadas'], barmode='group', title='Registros de entregas por curso')
        st.plotly_chart(fig, use_container_width=True)

with tab_diag:
    st.markdown('### Diagnóstico de acceso y registros Canvas')
    if errors:
        st.error('Algunos cursos no pudieron procesarse completamente:')
        st.dataframe(pd.DataFrame(errors), use_container_width=True, hide_index=True)
    st.markdown('#### Estudiantes detectados, activos, pendientes o no registrados')
    st.dataframe(diagnostic_df, use_container_width=True, hide_index=True)
    if not diagnostic_df.empty:
        diag_count = diagnostic_df.groupby(['curso_general','clasificacion_registro']).size().reset_index(name='cantidad') if 'clasificacion_registro' in diagnostic_df.columns else pd.DataFrame()
        if not diag_count.empty:
            fig = px.bar(diag_count, x='curso_general', y='cantidad', color='clasificacion_registro', title='Clasificación de registros detectados en Canvas')
            st.plotly_chart(fig, use_container_width=True)
    st.markdown('#### Criterios aplicados')
    st.markdown(f'''
- La app consulta cursos mediante token Canvas y consolida los cursos/secciones seleccionados.
- El agrupamiento general se realiza limpiando nombres con patrón `SECCIÓN`.
- Se incluyen estudiantes activos y también estudiantes detectados por Canvas que estén pendientes, inactivos, concluidos o con datos incompletos.
- Un estudiante con baja conexión no se prioriza como crítico si tiene entregables finalizados o al día.
- Umbral de finalización configurado: **{completion_threshold}%** del curso.
- Umbral para estudiante al día: **{current_threshold}%** de actividades actuales y cero pendientes/atrasadas.
''')

with tab_rep:
    st.markdown('### Exportables institucionales')
    # Se inserta la hoja resumen global dentro de la función existente usando argumentos actuales.
    xlsx = excel_bytes(summary, sub_df, pd.DataFrame(), pd.DataFrame(), modules_df, module_items_df, module_matrix)
    st.download_button('Descargar Excel global completo', xlsx, file_name=f'reporte_global_ave_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
    pdf = pdf_bytes(summary, 'Informe global AVE', f'{analysis.get("selected_count", 0)} cursos/secciones Canvas', generated_by, str(analysis_date))
    st.download_button('Descargar PDF ejecutivo global', pdf, file_name=f'informe_ejecutivo_global_ave_{datetime.now().strftime("%Y%m%d_%H%M")}.pdf', mime='application/pdf', use_container_width=True)
    st.markdown('#### Ficha individual PDF')
    active_options = {f"{r.get('estudiante')} | {r.get('curso_general')} | {r.get('correo','')}": r for _, r in summary.iterrows()}
    if active_options:
        key = st.selectbox('Seleccione estudiante', list(active_options.keys()))
        row = active_options[key].to_dict() if hasattr(active_options[key], 'to_dict') else dict(active_options[key])
        indiv_pdf = individual_pdf_bytes(row, sub_df, pd.DataFrame(), row.get('curso_general','Curso'), row.get('curso_canvas','Canvas'), generated_by, str(analysis_date))
        clean_name = ''.join(ch for ch in str(row.get('estudiante','estudiante')) if ch.isalnum() or ch in (' ','_','-')).strip().replace(' ','_')[:60]
        st.download_button('Descargar ficha individual PDF', indiv_pdf, file_name=f'ficha_{clean_name}.pdf', mime='application/pdf', use_container_width=True)

st.markdown(f'<div class="footer">Desarrollador: {DEV} | Universidad del Valle de Guatemala - AVE</div>', unsafe_allow_html=True)
