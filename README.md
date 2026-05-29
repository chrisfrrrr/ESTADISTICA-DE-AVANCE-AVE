# AVE Monitor Académico Pro 4.0 FIX9

Aplicación Streamlit para asesores AVE/UVG conectada directamente a Canvas mediante token.

## Funciones principales

- Conexión a Canvas por URL y token personal.
- Carga automática de cursos a los que el usuario tiene acceso.
- Selección de múltiples cursos/secciones Canvas.
- Agrupación automática por curso general, por ejemplo: Matemáticas, Fundamentos de la Comunicación e Iniciando tu Experiencia Virtual.
- Consulta de estudiantes activos.
- Inclusión de estudiantes detectados por Canvas que estén pendientes, inactivos, concluidos o con datos incompletos/no registrados.
- Consulta de actividades, entregas y módulos.
- Riesgo ajustado por entregables: un estudiante con baja conexión puede dejar de priorizarse si ya finalizó o está al día.
- Dashboard interno con gráficas.
- Exportación a Excel global.
- Exportación a PDF ejecutivo institucional AVE/UVG con gráficas.
- Ficha individual PDF por estudiante.

## Instalación local

```powershell
cd "RUTA\canvas_asesor_ave_PRO_4_0"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

## Uso recomendado

1. Ingresar URL de Canvas.
2. Ingresar token de acceso.
3. Presionar **Probar conexión / cargar cursos**.
4. Seleccionar todos los cursos o secciones Canvas que se desean consolidar.
5. Configurar fechas, horas mínimas y criterios de finalización.
6. Presionar **Generar análisis global**.
7. Revisar las pestañas de resumen, cursos globales, riesgo, estudiantes, entregables y diagnóstico.
8. Descargar Excel y PDF ejecutivo.

## Seguridad

El token no debe compartirse ni guardarse en documentos públicos. La app no necesita que el token esté escrito dentro del código.


## FIX9
Esta versión conserva el PDF ejecutivo premium de FIX4, pero restaura el cliente de conexión Canvas estable de FIX3 para evitar el error 406 posterior a FIX4.
