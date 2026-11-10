import streamlit as st

st.set_page_config(
    page_title="PIM - Demostracion web",
    page_icon="🛠️",
    layout="wide",
)

st.title("PIM - Demostracion web")
st.caption("Prueba de viabilidad de una aplicacion Python en navegador")

st.success("La aplicacion Python esta funcionando desde una interfaz web.")

st.subheader("Informacion de la prueba")
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Entorno", "Web")
with col2:
    st.metric("Tecnologia", "Python")
with col3:
    st.metric("Estado", "Operativo")

st.info(
    "Esta demostracion confirma que una aplicacion desarrollada en Python "
    "puede publicarse y abrirse desde un navegador. No contiene datos reales."
)

st.divider()
st.subheader("Prueba funcional")

equipo = st.text_input("Codigo de equipo", placeholder="Ejemplo: PALA-01")
archivo = st.file_uploader("Subir archivo de prueba", type=["xlsx", "xls", "csv"])

if st.button("Ejecutar prueba", type="primary"):
    if not equipo.strip():
        st.warning("Ingrese el codigo del equipo.")
    elif archivo is None:
        st.warning("Seleccione un archivo Excel o CSV.")
    else:
        st.success(f"Archivo recibido correctamente para {equipo}: {archivo.name}")
