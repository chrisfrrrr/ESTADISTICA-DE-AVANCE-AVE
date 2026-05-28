from __future__ import annotations
from datetime import datetime, date
import pandas as pd

RISK_SCORE = {'Bajo': 1, 'Medio': 2, 'Alto': 3}


def _to_dt(series):
    return pd.to_datetime(series, errors='coerce')


def build_pro3_tracking(summary: pd.DataFrame, history: pd.DataFrame | None, followups: pd.DataFrame | None, analysis_date=None) -> pd.DataFrame:
    """Enriquece el resumen actual con trazabilidad Pro 3.0.

    Agrega tendencia vs corte anterior, reincidencia de riesgo alto, estado de seguimiento,
    último contacto y acción recomendada ampliada para el asesor.
    """
    if summary is None or summary.empty:
        return pd.DataFrame()
    out = summary.copy()
    out['user_id_str'] = out['user_id'].astype(str)

    # Comparación contra corte anterior.
    out['riesgo_anterior'] = ''
    out['puntaje_anterior'] = None
    out['tendencia_riesgo'] = 'Sin histórico previo'
    out['reincidencia_alto'] = 0
    out['cortes_en_riesgo'] = 0
    out['clasificacion_reincidencia'] = 'Sin histórico suficiente'

    if history is not None and not history.empty and 'created_at' in history.columns:
        hist = history.copy()
        hist['created_at_dt'] = _to_dt(hist['created_at'])
        hist['user_id_str'] = hist['user_id'].astype(str)
        distinct = hist.dropna(subset=['created_at_dt'])['created_at_dt'].drop_duplicates().sort_values(ascending=False).tolist()
        prev_dt = distinct[1] if len(distinct) >= 2 else None
        if prev_dt is not None:
            prev = hist[hist['created_at_dt'].eq(prev_dt)].copy()
            prev = prev.sort_values('created_at_dt').drop_duplicates('user_id_str', keep='last')
            prev_map = prev.set_index('user_id_str')
            out['riesgo_anterior'] = out['user_id_str'].map(prev_map.get('riesgo_integral', pd.Series(dtype=object))) if 'riesgo_integral' in prev_map else ''
            out['puntaje_anterior'] = out['user_id_str'].map(prev_map.get('puntaje_riesgo', pd.Series(dtype=float))) if 'puntaje_riesgo' in prev_map else None
            def tendencia(row):
                ant = row.get('riesgo_anterior')
                act = row.get('riesgo_integral')
                if not ant or pd.isna(ant):
                    return 'Nuevo en seguimiento'
                if ant == act:
                    return 'Sin cambio'
                if RISK_SCORE.get(act, 0) > RISK_SCORE.get(ant, 0):
                    return 'Empeoró'
                if ant in ['Alto','Medio'] and act == 'Bajo':
                    return 'Recuperado'
                return 'Mejoró'
            out['tendencia_riesgo'] = out.apply(tendencia, axis=1)

        if 'riesgo_integral' in hist.columns:
            risk_hist = hist[hist['riesgo_integral'].isin(['Alto','Medio'])]
            high_hist = hist[hist['riesgo_integral'].eq('Alto')]
            out['reincidencia_alto'] = out['user_id_str'].map(high_hist.groupby('user_id_str').size()).fillna(0).astype(int)
            out['cortes_en_riesgo'] = out['user_id_str'].map(risk_hist.groupby('user_id_str').size()).fillna(0).astype(int)
            def rec_label(n):
                n = int(n or 0)
                if n >= 4: return 'Riesgo crítico sostenido'
                if n >= 2: return 'Riesgo recurrente'
                if n == 1: return 'Riesgo nuevo'
                return 'Sin reincidencia alta'
            out['clasificacion_reincidencia'] = out['reincidencia_alto'].apply(rec_label)

    # Seguimiento/bitácora.
    out['ultimo_contacto'] = ''
    out['resultado_ultimo_contacto'] = ''
    out['dias_desde_ultimo_contacto'] = None
    out['seguimiento_estado'] = 'Sin contacto registrado'
    if followups is not None and not followups.empty and 'created_at' in followups.columns:
        fu = followups.copy()
        fu['created_at_dt'] = _to_dt(fu['created_at'])
        fu['user_id_str'] = fu['user_id'].astype(str)
        latest = fu.sort_values('created_at_dt').drop_duplicates('user_id_str', keep='last').set_index('user_id_str')
        out['ultimo_contacto'] = out['user_id_str'].map(latest.get('created_at', pd.Series(dtype=object))).fillna('')
        out['resultado_ultimo_contacto'] = out['user_id_str'].map(latest.get('resultado', pd.Series(dtype=object))).fillna('')
        out['fecha_proxima_accion'] = out['user_id_str'].map(latest.get('fecha_proxima_accion', pd.Series(dtype=object))).fillna('')
        last_dt = out['user_id_str'].map(latest['created_at_dt']) if 'created_at_dt' in latest.columns else pd.Series([pd.NaT]*len(out))
        today = pd.Timestamp(analysis_date or date.today())
        out['dias_desde_ultimo_contacto'] = (today - pd.to_datetime(last_dt, errors='coerce').dt.tz_localize(None)).dt.days
        def seg_estado(row):
            if not row.get('ultimo_contacto'):
                return 'Sin contacto registrado'
            prox = pd.to_datetime(row.get('fecha_proxima_accion'), errors='coerce')
            if pd.notna(prox) and prox.date() < (analysis_date or date.today()):
                return 'Seguimiento vencido'
            if row.get('resultado_ultimo_contacto') in ['No respondió']:
                return 'Contactado sin respuesta'
            if row.get('resultado_ultimo_contacto') in ['Respondió','Se comprometió','Contactado']:
                return 'Con seguimiento activo'
            return 'Con contacto registrado'
        out['seguimiento_estado'] = out.apply(seg_estado, axis=1)

    def accion30(row):
        riesgo = row.get('riesgo_integral')
        tend = row.get('tendencia_riesgo')
        seg = row.get('seguimiento_estado')
        reinc = row.get('clasificacion_reincidencia')
        if riesgo == 'Alto' and seg == 'Sin contacto registrado':
            return 'Enviar mensaje inicial y registrar bitácora hoy'
        if riesgo == 'Alto' and seg in ['Contactado sin respuesta','Seguimiento vencido']:
            return 'Recontactar y considerar escalamiento a coordinación'
        if reinc == 'Riesgo crítico sostenido':
            return 'Escalar caso y documentar plan de intervención'
        if tend == 'Empeoró':
            return 'Priorizar revisión; pasó a un nivel de riesgo mayor'
        if tend == 'Recuperado':
            return 'Enviar refuerzo positivo y mantener monitoreo'
        if riesgo == 'Medio':
            return 'Enviar recordatorio preventivo y revisar próximos pendientes'
        return row.get('accion_recomendada', 'Monitoreo regular')
    out['accion_pro_3'] = out.apply(accion30, axis=1)
    return out.drop(columns=['user_id_str'], errors='ignore')


def build_course_control_kpis(enhanced: pd.DataFrame) -> dict:
    if enhanced is None or enhanced.empty:
        return {}
    return {
        'nuevos_alto': int(((enhanced.get('riesgo_integral') == 'Alto') & (enhanced.get('tendencia_riesgo').isin(['Empeoró','Nuevo en seguimiento','Sin histórico previo']))).sum()),
        'recuperados': int((enhanced.get('tendencia_riesgo') == 'Recuperado').sum()),
        'empeoraron': int((enhanced.get('tendencia_riesgo') == 'Empeoró').sum()),
        'reincidentes': int(enhanced.get('clasificacion_reincidencia', pd.Series(dtype=str)).isin(['Riesgo recurrente','Riesgo crítico sostenido']).sum()),
        'seguimientos_vencidos': int((enhanced.get('seguimiento_estado') == 'Seguimiento vencido').sum()),
        'sin_contacto_alto': int(((enhanced.get('riesgo_integral') == 'Alto') & (enhanced.get('seguimiento_estado') == 'Sin contacto registrado')).sum()),
    }


def module_heatmap_summary(module_matrix: pd.DataFrame | None) -> pd.DataFrame:
    if module_matrix is None or module_matrix.empty:
        return pd.DataFrame()
    m = module_matrix.copy()
    if 'estado' not in m.columns:
        return pd.DataFrame()
    grouped = m.groupby(['modulo','estado']).size().reset_index(name='cantidad')
    total = m.groupby('modulo').size().reset_index(name='total_registros')
    pivot = grouped.pivot_table(index='modulo', columns='estado', values='cantidad', aggfunc='sum', fill_value=0).reset_index()
    out = pivot.merge(total, on='modulo', how='left')
    for col in ['Entregado','Pendiente actual','Pendiente futuro','Atrasado','No aplica']:
        if col not in out.columns:
            out[col] = 0
    out['pendientes_actuales_modulo'] = out['Pendiente actual'] + out['Atrasado']
    out['cumplimiento_modulo_%'] = (out['Entregado'] / out['total_registros'].replace(0, pd.NA) * 100).fillna(0).round(1)
    return out.sort_values(['pendientes_actuales_modulo','Atrasado'], ascending=False)


def data_quality_summary(assign_df: pd.DataFrame | None, module_items_df: pd.DataFrame | None, sub_df: pd.DataFrame | None) -> pd.DataFrame:
    rows = []
    def add(ind, valor, impacto):
        rows.append({'control': ind, 'valor': valor, 'impacto': impacto})
    assign = assign_df if assign_df is not None else pd.DataFrame()
    items = module_items_df if module_items_df is not None else pd.DataFrame()
    subs = sub_df if sub_df is not None else pd.DataFrame()
    add('Actividades detectadas por API', len(assign), 'Base para cálculo de entregas y avance')
    add('Ítems de módulos detectados', len(items), 'Base para vista por módulos')
    if not assign.empty and 'fecha_entrega' in assign.columns:
        add('Actividades sin fecha de entrega', int(assign['fecha_entrega'].isna().sum()), 'Se consideran actuales para seguimiento')
    if not items.empty and 'es_entregable' in items.columns:
        add('Ítems marcados como entregables', int(items['es_entregable'].fillna(False).sum()), 'Se cruzan contra submissions de estudiantes')
    if not items.empty and not assign.empty and 'content_id' in items.columns and 'assignment_id' in assign.columns:
        linked = items['content_id'].isin(assign['assignment_id']).sum()
        add('Ítems de módulo vinculados a actividades', int(linked), 'Permite matriz estudiante vs entregable')
    add('Submissions recibidas por API', len(subs), 'Base para entregado/pendiente/atrasado')
    return pd.DataFrame(rows)
