import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import openpyxl
from io import BytesIO

# ---------------------- Configuración de Autenticación ----------------------
# Se utiliza la información de secrets.toml para autenticar con Google
creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"])
client = gspread.authorize(creds)

# ID de la hoja de Google Sheets (asegúrate de que el enlace compartido corresponda)
SHEET_ID = "1srAGigOz4fI9tfYTAP1-ens9M27-1TapSQaLZIgEhDE"
gsheet = client.open_by_key(SHEET_ID)

# ---------------------- Funciones de la Aplicación ----------------------

def ingresar_titular():
    st.header("Ingresar Cuenta Titular")
    rut = st.text_input("Rut Titular (Ej. 12345678-9)")
    nombre = st.text_input("Nombre del Titular")
    email = st.text_input("Email del Titular")
    cuenta = st.text_input("Número de Cuenta Titular")
    
    if st.button("Guardar Titular"):
        try:
            # Se accede a la hoja "Ctas" y se agrega una nueva fila
            ws = gsheet.worksheet("Ctas")
            data = ws.get_all_values()
            row_index = len(data) + 1  # Se asume que la primera fila es encabezado
            # Se utiliza "1" en la columna A para identificar un titular
            ws.update(f"A{row_index}", [[1, rut, nombre, email, 39, cuenta, "Cuenta Corriente"]])
            
            # Actualizar la hoja "clientes" con el Rut y nombre del titular
            ws_clientes = gsheet.worksheet("clientes")
            clientes_data = ws_clientes.get_all_values()
            if not any(row and row[0] == rut for row in clientes_data):
                ws_clientes.append_row([rut, nombre])
            
            st.success("Titular guardado correctamente.")
        except Exception as e:
            st.error(f"Error: {e}")

def ingresar_destinatario():
    st.header("Ingresar Cuenta Destinatario")
    rut = st.text_input("Rut Destinatario (Ej. 12345678-9)", key="dest_rut")
    nombre = st.text_input("Nombre del Destinatario", key="dest_nombre")
    email = st.text_input("Email del Destinatario", key="dest_email")
    banco_options = [
        "Banco Chile", "Banco Internacional", "Banco Estado", "Scotiabank",
        "Banco BCI", "CorpBanca", "Banco Bice", "HSBC Bank", "Santander",
        "ITAU CorpBanca", "Banco Security", "Banco Falabella", "Banco Ripley",
        "Banco Consorcio", "Banco Paris", "ScotiaBank Azul", "B. Desarrollo",
        "Coopeuch", "Prepago Los Héroes", "Tenpo Prepago"
    ]
    banco = st.selectbox("Seleccione Banco", banco_options, key="dest_banco")
    cuenta = st.text_input("Número de Cuenta Destinatario", key="dest_cuenta")
    tipo_cuenta = st.selectbox("Tipo de Cuenta", ["Cuenta Corriente", "Cuenta de Ahorro", "Cuenta Vista"], key="dest_tipo")
    
    if st.button("Guardar Destinatario"):
        try:
            ws = gsheet.worksheet("Ctas")
            data = ws.get_all_values()
            row_index = len(data) + 1
            # Se utiliza "2" en la columna A para identificar un destinatario
            ws.update(f"A{row_index}", [[2, rut, nombre, email, "", "", "", banco, cuenta, tipo_cuenta]])
            
            # Actualizar la hoja "clientes"
            ws_clientes = gsheet.worksheet("clientes")
            clientes_data = ws_clientes.get_all_values()
            if not any(row and row[0] == rut for row in clientes_data):
                ws_clientes.append_row([rut, nombre])
            
            st.success("Destinatario guardado correctamente.")
        except Exception as e:
            st.error(f"Error: {e}")

def ingresar_nomina():
    st.header("Ingresar Nómina")
    try:
        ws = gsheet.worksheet("Ctas")
        data = ws.get_all_records()
        # Se filtran registros: "1" para titulares y "2" para destinatarios.
        titulares = [row for row in data if row.get("A") == 1 or row.get("ID") == 1]
        destinatarios = [row for row in data if row.get("A") == 2 or row.get("ID") == 2]
    except Exception as e:
        st.error(f"Error al obtener datos: {e}")
        titulares = []
        destinatarios = []
    
    if not titulares or not destinatarios:
        st.warning("No se encontraron titulares o destinatarios en la base de datos.")
        return

    # Se extraen los nombres para mostrarlos en los selectboxes
    titular_names = [t["Nombre"] for t in titulares if "Nombre" in t]
    destinatario_names = [d["Nombre"] for d in destinatarios if "Nombre" in d]
    
    titular_sel = st.selectbox("Seleccione Cuenta Titular", titular_names)
    destinatario_sel = st.selectbox("Seleccione Destinatario", destinatario_names)
    monto = st.number_input("Monto a Transferir", min_value=0, step=1000)
    glosa = st.text_input("Glosa")
    
    if st.button("Procesar Nómina"):
        try:
            # Se carga la plantilla de Excel (debe estar en la raíz del repositorio)
            wb = openpyxl.load_workbook("Plantilla_Transferencia.xlsx")
            ws_excel = wb.active
            
            # Actualizar celdas con los datos ingresados
            ws_excel["B3"].value = titular_sel.replace("-", "")
            ws_excel["B9"].value = 12345678  # Este valor idealmente se debe obtener del registro del titular
            ws_excel["B10"].value = glosa
            ws_excel["B11"].value = glosa

            # Dividir el monto en partes de hasta 7.000.000
            montos = []
            temp = monto
            while temp > 7000000:
                montos.append(7000000)
                temp -= 7000000
            if temp > 0:
                montos.append(temp)
            
            fila_inicio = 15
            for i, parte in enumerate(montos):
                fila = fila_inicio + i
                ws_excel[f"A{fila}"].value = destinatario_sel.replace("-", "")
                ws_excel[f"B{fila}"].value = destinatario_sel
                ws_excel[f"C{fila}"].value = parte
                ws_excel[f"D{fila}"].value = "Abono en Cuenta"
                # Se actualizan las columnas para la glosa (ajusta según la estructura de tu plantilla)
                ws_excel[f"I{fila}"].value = glosa
                ws_excel[f"J{fila}"].value = glosa
                ws_excel[f"K{fila}"].value = glosa
                ws_excel[f"L{fila}"].value = glosa

            # Guardar el archivo en memoria para ofrecer la descarga
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            st.download_button(
                "Descargar Nómina",
                data=buffer,
                file_name="Nomina.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Nómina procesada correctamente.")
        except Exception as e:
            st.error(f"Error al procesar nómina: {e}")

# ---------------------- Menú de Navegación ----------------------
st.sidebar.title("Menú Principal")
opcion = st.sidebar.radio("Seleccione una opción:", (
    "Ingresar Cuenta Titular", 
    "Ingresar Cuenta Destinatario", 
    "Ingresar Nómina"
))

if opcion == "Ingresar Cuenta Titular":
    ingresar_titular()
elif opcion == "Ingresar Cuenta Destinatario":
    ingresar_destinatario()
elif opcion == "Ingresar Nómina":
    ingresar_nomina()
