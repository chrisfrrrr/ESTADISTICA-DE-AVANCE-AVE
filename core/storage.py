from __future__ import annotations
import sqlite3
from pathlib import Path
from datetime import datetime
import pandas as pd

DB_PATH = Path('data/historial_ave_canvas.db')

SNAPSHOT_COLS = [
    'created_at','course_id','course_name','section_id','section_name','user_id','estudiante','nombre','correo','ultima_actividad',
    'horas_sin_actividad','riesgo_desconexion','tiempo_total_horas','horas_esperadas','deficit_horas','cumplimiento_horas',
    'actividades_total','entregadas','pendientes','atrasadas','porcentaje_avance','promedio_score','puntaje_riesgo','riesgo_integral','segmento_ave','accion_recomendada'
]

FOLLOWUP_COLS = [
    'id','created_at','course_id','course_name','section_id','section_name','user_id','estudiante','correo',
    'medio','motivo','resultado','proxima_accion','fecha_proxima_accion','observaciones','registrado_por'
]

def _connect():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    return sqlite3.connect(DB_PATH)

def init_db():
    with _connect() as con:
        con.execute('''CREATE TABLE IF NOT EXISTS snapshots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            course_id TEXT,
            course_name TEXT,
            section_id TEXT,
            section_name TEXT,
            user_id TEXT,
            estudiante TEXT,
            nombre TEXT,
            correo TEXT,
            ultima_actividad TEXT,
            horas_sin_actividad REAL,
            riesgo_desconexion TEXT,
            tiempo_total_horas REAL,
            horas_esperadas REAL,
            deficit_horas REAL,
            cumplimiento_horas TEXT,
            actividades_total REAL,
            entregadas REAL,
            pendientes REAL,
            atrasadas REAL,
            porcentaje_avance REAL,
            promedio_score REAL,
            puntaje_riesgo REAL,
            riesgo_integral TEXT,
            segmento_ave TEXT,
            accion_recomendada TEXT
        )''')
        con.execute('''CREATE TABLE IF NOT EXISTS followups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            course_id TEXT,
            course_name TEXT,
            section_id TEXT,
            section_name TEXT,
            user_id TEXT,
            estudiante TEXT,
            correo TEXT,
            medio TEXT,
            motivo TEXT,
            resultado TEXT,
            proxima_accion TEXT,
            fecha_proxima_accion TEXT,
            observaciones TEXT,
            registrado_por TEXT
        )''')
        con.commit()

def save_snapshot(df: pd.DataFrame, course_id, course_name, section_id, section_name, created_at: str):
    init_db()
    if df is None or df.empty:
        return
    temp = df.copy()
    temp['created_at'] = created_at
    temp['course_id'] = str(course_id)
    temp['course_name'] = course_name
    temp['section_id'] = '' if section_id is None else str(section_id)
    temp['section_name'] = section_name or 'Todas'
    for c in SNAPSHOT_COLS:
        if c not in temp.columns:
            temp[c] = None
    temp['ultima_actividad'] = temp['ultima_actividad'].astype(str)
    with _connect() as con:
        temp[SNAPSHOT_COLS].to_sql('snapshots', con, if_exists='append', index=False)

def load_history(course_id=None, section_id=None) -> pd.DataFrame:
    init_db()
    with _connect() as con:
        q = 'SELECT * FROM snapshots'
        params = []
        clauses = []
        if course_id:
            clauses.append('course_id=?')
            params.append(str(course_id))
        if section_id is not None:
            clauses.append('section_id=?')
            params.append(str(section_id))
        if clauses:
            q += ' WHERE ' + ' AND '.join(clauses)
        q += ' ORDER BY created_at DESC'
        return pd.read_sql_query(q, con, params=params)

def save_followup(course_id, course_name, section_id, section_name, user_id, estudiante, correo, medio, motivo, resultado, proxima_accion, fecha_proxima_accion, observaciones, registrado_por):
    init_db()
    row = {
        'created_at': datetime.now().isoformat(timespec='seconds'),
        'course_id': str(course_id),
        'course_name': course_name,
        'section_id': '' if section_id is None else str(section_id),
        'section_name': section_name or 'Todas',
        'user_id': str(user_id),
        'estudiante': estudiante,
        'correo': correo,
        'medio': medio,
        'motivo': motivo,
        'resultado': resultado,
        'proxima_accion': proxima_accion,
        'fecha_proxima_accion': str(fecha_proxima_accion) if fecha_proxima_accion else '',
        'observaciones': observaciones,
        'registrado_por': registrado_por,
    }
    with _connect() as con:
        pd.DataFrame([row]).to_sql('followups', con, if_exists='append', index=False)

def load_followups(course_id=None, section_id=None, user_id=None) -> pd.DataFrame:
    init_db()
    with _connect() as con:
        q = 'SELECT * FROM followups'
        params = []
        clauses = []
        if course_id:
            clauses.append('course_id=?')
            params.append(str(course_id))
        if section_id is not None:
            clauses.append('section_id=?')
            params.append(str(section_id))
        if user_id:
            clauses.append('user_id=?')
            params.append(str(user_id))
        if clauses:
            q += ' WHERE ' + ' AND '.join(clauses)
        q += ' ORDER BY created_at DESC'
        return pd.read_sql_query(q, con, params=params)

def delete_followup(followup_id: int):
    init_db()
    with _connect() as con:
        con.execute('DELETE FROM followups WHERE id=?', (int(followup_id),))
        con.commit()
