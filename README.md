# INMAC · Planilla de fichajes

## Archivos del repositorio
- `app.py`: interfaz web en Streamlit.
- `planilla_fichajes_universal_legajo_v2.py`: motor principal.
- `requirements.txt`: dependencias.
- `logo_inmac.jpg`: logo de la empresa.
- `.gitignore`: exclusiones recomendadas.

## Cómo publicarlo
1. Crear un repositorio nuevo en GitHub.
2. Subir estos archivos al repo.
3. Entrar a https://share.streamlit.io/
4. Elegir el repositorio y publicar usando `app.py` como archivo principal.

## Qué hace
- Detecta automáticamente cuál archivo es fichajes y cuál es planilla base.
- Identifica operarios por nombre y por legajo.
- Adapta encabezados y fórmulas diarias según el mes y el día de la semana.
- Genera el Excel final listo para descargar.
