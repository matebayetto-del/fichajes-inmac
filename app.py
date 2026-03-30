from __future__ import annotations

import io
import os
import zipfile
import tempfile
from pathlib import Path
from contextlib import contextmanager

import streamlit as st

import planilla_fichajes_universal_legajo_v2 as motor


st.set_page_config(
    page_title="INMAC | Planilla de fichajes",
    page_icon="📋",
    layout="centered",
)


@contextmanager
def temp_chdir(path: Path):
    anterior = Path.cwd()
    os.chdir(path)
    try:
        yield
    finally:
        os.chdir(anterior)


def guardar_upload(uploaded_file, destino: Path) -> Path:
    ruta = destino / uploaded_file.name
    ruta.write_bytes(uploaded_file.getbuffer())
    return ruta


def resumen_diagnostico(diagnostico: list[dict]) -> list[dict]:
    salida = []
    for item in diagnostico:
        salida.append(
            {
                "Archivo": Path(item["path"]).name,
                "Score planilla": item.get("score_planilla"),
                "Score fichajes": item.get("score_fichajes"),
            }
        )
    return salida


def procesar_desde_streamlit(uploads: list, usar_feriados_argentina: bool = True):
    if len(uploads) < 2:
        raise ValueError("Tenés que subir los 2 archivos Excel.")

    with tempfile.TemporaryDirectory() as tmpdir:
        carpeta = Path(tmpdir)
        rutas = [guardar_upload(u, carpeta) for u in uploads]

        archivo_planilla, archivo_fichajes, diagnostico = motor.detectar_archivos([str(p) for p in rutas])

        if usar_feriados_argentina:
            motor.PAIS_FERIADOS = "AR"
            motor.SUBDIV_FERIADOS = None

        with temp_chdir(carpeta):
            nombres_salida = motor.procesar_archivos(
                archivo_planilla=archivo_planilla,
                archivo_fichajes=archivo_fichajes,
                descargar_en_colab=False,
            )

        salidas = [carpeta / nombre for nombre in nombres_salida]
        blobs = {archivo.name: archivo.read_bytes() for archivo in salidas}

        return {
            "archivo_planilla": Path(archivo_planilla).name,
            "archivo_fichajes": Path(archivo_fichajes).name,
            "diagnostico": resumen_diagnostico(diagnostico),
            "salidas": blobs,
        }


def mostrar_header():
    logo_path = Path(__file__).with_name("logo_inmac.jpg")
    col_logo, col_texto = st.columns([1, 3])
    with col_logo:
        if logo_path.exists():
            st.image(str(logo_path), use_container_width=True)
    with col_texto:
        st.title("Planilla de fichajes y horas")
        st.caption("INMAC · Subí los 2 Excel de la obra, procesá y descargá el archivo final.")


mostrar_header()

with st.container(border=True):
    st.subheader("Cómo usarlo")
    st.markdown(
        "1. Subí el archivo de fichadas y la planilla base.\n"
        "2. Tocá **Procesar archivos**.\n"
        "3. Descargá el Excel final generado por la app."
    )

with st.container(border=True):
    st.subheader("Cargar archivos")
    archivos = st.file_uploader(
        "Arrastrá o seleccioná los 2 Excel",
        type=["xls", "xlsx", "xlsm", "xltx", "xltm", "xlsb"],
        accept_multiple_files=True,
        help=(
            "No hace falta que tengan nombres específicos. La app intenta detectar cuál es planilla base y cuál es fichajes."
        ),
    )

    if archivos:
        st.write("**Archivos cargados:**")
        for archivo in archivos:
            st.write(f"- {archivo.name}")

    col1, col2 = st.columns(2)
    with col1:
        usar_feriados_argentina = st.checkbox("Usar feriados de Argentina", value=True)
    with col2:
        st.write("")
        procesar = st.button("Procesar archivos", type="primary", use_container_width=True)

if procesar:
    try:
        if not archivos or len(archivos) < 2:
            st.error("Subí ambos archivos antes de procesar.")
        else:
            with st.spinner("Procesando planillas..."):
                resultado = procesar_desde_streamlit(archivos, usar_feriados_argentina)

            st.success("Proceso completado.")

            with st.container(border=True):
                st.subheader("Archivos detectados")
                st.write(f"**Planilla base:** {resultado['archivo_planilla']}")
                st.write(f"**Fichajes:** {resultado['archivo_fichajes']}")
                st.dataframe(resultado["diagnostico"], use_container_width=True, hide_index=True)

            with st.container(border=True):
                st.subheader("Descargas")
                salidas = resultado["salidas"]
                nombres = list(salidas.keys())

                if len(nombres) == 1:
                    nombre = nombres[0]
                    st.download_button(
                        label=f"Descargar {nombre}",
                        data=salidas[nombre],
                        file_name=nombre,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                else:
                    zip_bytes = io.BytesIO()
                    with zipfile.ZipFile(zip_bytes, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                        for nombre, contenido in salidas.items():
                            zf.writestr(nombre, contenido)
                    zip_bytes.seek(0)

                    st.download_button(
                        label="Descargar resultados (.zip)",
                        data=zip_bytes.getvalue(),
                        file_name="planillas_procesadas.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )

                    st.write("**Archivos generados:**")
                    for nombre in nombres:
                        st.write(f"- {nombre}")

            with st.expander("Consejos si algo no coincide"):
                st.write(
                    "Si algún operario no se identifica por nombre, el motor también intenta emparejar por legajo. "
                    "En la salida quedan hojas auxiliares de revisión con coincidencias dudosas e incidencias."
                )

    except Exception as e:
        st.error(f"No se pudo completar el proceso: {e}")

with st.expander("Qué hace esta app"):
    st.write(
        "Detecta automáticamente cuál Excel es la planilla base y cuál corresponde a fichajes, "
        "ubica el mes desde las fechas cargadas, adapta la planilla al calendario real y genera el archivo final listo para descargar."
    )
