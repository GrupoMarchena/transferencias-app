import streamlit as st
import pandas as pd
import openpyxl
import tempfile
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from io import BytesIO

# === CARGAR DATOS DESDE GOOGLE SHEETS ===
@st.cache_data
def cargar_datos():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_dict = st.secrets["google_service_account"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(json.loads(json.dumps(creds_dict)), scope)
    client = gspread.authorize(creds)

    sheet = client.open_by_key("1srAGigOz4fI9tfYTAP1-ens9M27-1TapSQaLZIgEhDE")
    clientes_ws = sheet.worksheet("clientes")
    ctas_ws = sheet.worksheet("ctas")

    clientes = pd.DataFrame(clientes_ws.get_all_records())
    ctas = pd.DataFrame(ctas_ws.get_all_records())

    return clientes, ctas, sheet

clientes, ctas, hoja = cargar_datos()

st.set_page_config(page_title="Nómina de Transferencias", page_icon="💸", layout="centered")
st.title("💸 Aplicación de Nómina de Transferencias")
tabs = st.tabs(["Ingresar Nómina", "Ingresar Titular", "Ingresar Proveedor"])

# === TAB 1: INGRESAR NÓMINA ===
with tabs[0]:
    st.subheader("📋 Generar Nómina de Transferencia")

    titulares = clientes[clientes["tipo"] == 1]
    proveedores = clientes[clientes["tipo"] == 2]

    with st.form("form_nomina"):
        titular = st.selectbox("Seleccionar Titular", titulares["nombre"].tolist())
        rut_titular = titulares[titulares["nombre"] == titular]["rut"].values[0]

        proveedor = st.selectbox("Seleccionar Proveedor", proveedores["nombre"].tolist())
        rut_prov = proveedores[proveedores["nombre"] == proveedor]["rut"].values[0]

        bancos_prov = ctas[ctas["rut_cliente"] == rut_prov]
        banco_sel = st.selectbox("Seleccionar Banco del Proveedor", bancos_prov["banco"].tolist())
        datos_banco = bancos_prov[bancos_prov["banco"] == banco_sel].iloc[0]

        monto = st.number_input("Monto a transferir", min_value=1000, step=1000)
        glosa = st.text_input("Glosa o detalle")

        generar = st.form_submit_button("Generar archivo Excel")

    if generar:
        URL_PLANTILLA = st.secrets["google_service_account"]["URL_PLANTILLA"]
        response = requests.get(URL_PLANTILLA)
        wb = openpyxl.load_workbook(BytesIO(response.content))
        ws = wb.active

        ws["B3"] = rut_titular
        ws["B9"] = datos_banco["cuenta"]
        ws["B10"] = glosa
        ws["B11"] = glosa

        fila = 15
        ws.cell(fila, 1, datos_banco["banco"])
        ws.cell(fila, 2, datos_banco["tipo_cuenta"])
        ws.cell(fila, 3, datos_banco["cuenta"])
        ws.cell(fila, 4, proveedor)
        ws.cell(fila, 5, rut_prov)
        ws.cell(fila, 6, monto)
        ws.cell(fila, 7, datos_banco["correo"])
        ws.cell(fila, 8, glosa)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            wb.save(tmp.name)
            with open(tmp.name, "rb") as f:
                st.success("✅ Archivo generado con éxito")
                st.download_button("📅 Descargar Excel", f, "transferencia.xlsx")

# === TAB 2: INGRESAR TITULAR ===
with tabs[1]:
    st.subheader("📄 Ingresar nueva cuenta titular")
    with st.form("form_titular"):
        nombre = st.text_input("Nombre titular")
        rut = st.text_input("RUT titular")
        guardar = st.form_submit_button("Guardar titular")

    if guardar:
        hoja.worksheet("clientes").append_row([rut, nombre, 1])
        st.success("Titular guardado exitosamente.")

# === TAB 3: INGRESAR PROVEEDOR ===
with tabs[2]:
    st.subheader("🛌 Ingresar proveedor y cuenta bancaria")
    with st.form("form_proveedor"):
        nombre = st.text_input("Nombre proveedor")
        rut = st.text_input("RUT proveedor")
        banco = st.text_input("Banco")
        tipo_cuenta = st.selectbox("Tipo de cuenta", ["Corriente", "Vista", "Ahorro"])
        cuenta = st.text_input("Nº de cuenta")
        correo = st.text_input("Correo electrónico")
        guardar = st.form_submit_button("Guardar proveedor")

    if guardar:
        hoja.worksheet("clientes").append_row([rut, nombre, 2])
        hoja.worksheet("ctas").append_row([rut, banco, tipo_cuenta, cuenta, correo])
        st.success("Proveedor y cuenta guardados.")

