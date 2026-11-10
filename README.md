# PIM Web Demo

Demostracion tecnica para comprobar que una aplicacion Python puede ejecutarse en una web y compartirse mediante una URL.

## Contenido

- `web_app.py`: interfaz web de demostracion.
- `requirements.txt`: dependencia necesaria para el despliegue.
- `.gitignore`: evita subir archivos locales o sensibles.

## Ejecutar localmente

1. Abrir una terminal dentro de esta carpeta.
2. Instalar la dependencia: `python -m pip install -r requirements.txt`
3. Ejecutar: `python -m streamlit run web_app.py`
4. Abrir `http://localhost:8501`

## Publicar en Streamlit Community Cloud

1. Crear un repositorio en GitHub.
2. Subir solamente los archivos de esta carpeta.
3. En Streamlit Community Cloud, seleccionar el repositorio, la rama `main` y `web_app.py`.
4. Pulsar **Deploy**.

Esta version no contiene la aplicacion PIM completa, la base SQLite ni informacion empresarial. Sirve para demostrar alojamiento y acceso web.
