import streamlit as st
import pandas as pd
import openpyxl
import io
from clean_excel import clean_card, clean_account, clean_dni, clean_phone, to_str_no_sci

# Page configuration
st.set_page_config(
    page_title="Limpiador de Base de Datos - Excel",
    page_icon="🧼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling for premium feel
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        border-radius: 8px;
        padding: 10px 24px;
        font-weight: bold;
        border: none;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #45a049;
        transform: scale(1.02);
    }
    .report-card {
        padding: 20px;
        border-radius: 10px;
        background-color: white;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 20px;
    }
    h1, h2, h3 {
        color: #2C3E50;
    }
    </style>
""", unsafe_allow_html=True)

# App Header
st.title("🧼 Limpiador y Cruzador de Bases de Datos Excel")
st.markdown("Carga tu archivo Excel de gestión de pendientes para validar campos y exportar las bases limpias.")

# File Uploader
uploaded_file = st.file_uploader("Selecciona el archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Load workbook in read-only to list sheet names quickly
        wb_check = openpyxl.load_workbook(uploaded_file, read_only=True)
        sheets = wb_check.sheetnames
        
        # User input / Selection of the sheet
        st.sidebar.header("Configuración de Hojas")
        sheet_name = st.sidebar.selectbox(
            "Selecciona la hoja a procesar",
            options=sheets,
            index=0
        )
        
        # Process Button
        if st.sidebar.button("Procesar Archivo"):
            with st.spinner("Procesando y validando base de datos..."):
                # Reset file pointer and read sheet
                uploaded_file.seek(0)
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                
                required_cols = ["TARJETA", "CUENTA", "DNI", "CLIENTE", "TELEFONO"]
                missing = [c for c in required_cols if c not in df.columns]
                
                if missing:
                    st.error(f"❌ Error: El archivo no contiene las columnas requeridas: {missing}")
                else:
                    alerts = []
                    clean_indices = []
                    
                    cleaned_tarjeta = []
                    cleaned_cuenta = []
                    cleaned_dni_list = []
                    cleaned_cliente = []
                    cleaned_telefono = []
                    
                    # Process and validate
                    for idx, row in df.iterrows():
                        raw_tarjeta = row["TARJETA"]
                        raw_cuenta = row["CUENTA"]
                        raw_dni = row["DNI"]
                        raw_cliente = row["CLIENTE"]
                        raw_telefono = row["TELEFONO"]
                        
                        c_tarjeta = clean_card(raw_tarjeta)
                        c_cuenta = clean_account(raw_cuenta)
                        c_dni = clean_dni(raw_dni)
                        c_cliente = str(raw_cliente).strip() if not pd.isna(raw_cliente) else ""
                        c_telefono = clean_phone(raw_telefono)
                        
                        failed_fields = []
                        if c_tarjeta is None:
                            failed_fields.append(f"TARJETA ({raw_tarjeta})")
                        if c_cuenta is None:
                            failed_fields.append(f"CUENTA ({raw_cuenta})")
                        if c_dni is None:
                            failed_fields.append(f"DNI ({raw_dni})")
                        if not c_cliente:
                            failed_fields.append(f"CLIENTE ({raw_cliente})")
                        if c_telefono is None:
                            failed_fields.append(f"TELEFONO ({raw_telefono})")
                            
                        if failed_fields:
                            dni_display = str(raw_dni).strip() if not pd.isna(raw_dni) else "SIN DNI"
                            if dni_display.endswith('.0'):
                                dni_display = dni_display[:-2]
                            alerts.append({
                                "Fila Excel": idx + 2,
                                "DNI": dni_display,
                                "Cliente": c_cliente or str(raw_cliente),
                                "Detalles": ", ".join(failed_fields)
                            })
                        else:
                            clean_indices.append(idx)
                            cleaned_tarjeta.append(c_tarjeta)
                            cleaned_cuenta.append(c_cuenta)
                            cleaned_dni_list.append(c_dni)
                            cleaned_cliente.append(c_cliente)
                            cleaned_telefono.append(c_telefono)
                            
                    # Filter and update dataframe for BASE PROLIJA
                    df_clean = df.loc[clean_indices].copy()
                    df_clean["TARJETA"] = cleaned_tarjeta
                    df_clean["CUENTA"] = cleaned_cuenta
                    df_clean["DNI"] = cleaned_dni_list
                    df_clean["CLIENTE"] = cleaned_cliente
                    df_clean["TELEFONO"] = cleaned_telefono
                    
                    for col in ["TARJETA", "CUENTA", "DNI", "TELEFONO"]:
                        df_clean[col] = df_clean[col].astype(str)
                        
                    # 1. Generate BASE PROLIJA in memory
                    prolija_buffer = io.BytesIO()
                    wb_prolija = openpyxl.Workbook()
                    ws_prolija = wb_prolija.active
                    ws_prolija.title = "Base Limpia"
                    
                    ws_prolija.append(list(df_clean.columns))
                    for _, row in df_clean.iterrows():
                        row_vals = []
                        for col in df_clean.columns:
                            val = row[col]
                            row_vals.append(to_str_no_sci(val))
                        ws_prolija.append(row_vals)
                        
                    for col_idx in range(1, len(df_clean.columns) + 1):
                        for row_idx in range(2, len(df_clean) + 2):
                            cell = ws_prolija.cell(row=row_idx, column=col_idx)
                            cell.number_format = '@'
                            
                    wb_prolija.save(prolija_buffer)
                    prolija_data = prolija_buffer.getvalue()
                    
                    # 2. Generate subir a predictivo in memory
                    df_pred = pd.DataFrame({
                        "TELEFONO": cleaned_telefono,
                        "CUENTA": cleaned_cuenta,
                        "DNI": cleaned_dni_list,
                        "TARJETA": cleaned_tarjeta,
                        "NOMBRE": cleaned_cliente
                    })
                    
                    pred_buffer = io.BytesIO()
                    wb_pred = openpyxl.Workbook()
                    ws_pred = wb_pred.active
                    ws_pred.title = "Predictivo"
                    
                    ws_pred.append(list(df_pred.columns))
                    for _, row in df_pred.iterrows():
                        row_vals = [to_str_no_sci(val) for val in row]
                        ws_pred.append(row_vals)
                        
                    for col_idx in range(1, len(df_pred.columns) + 1):
                        for row_idx in range(2, len(df_pred) + 2):
                            cell = ws_pred.cell(row=row_idx, column=col_idx)
                            cell.number_format = '@'
                            
                    wb_pred.save(pred_buffer)
                    pred_data = pred_buffer.getvalue()
                    
                    # UI Layout for Results
                    st.success("✅ ¡Procesamiento completado con éxito!")
                    
                    # Metrics
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Registros", len(df))
                    with col2:
                        st.metric("Registros Limpios (Cruzaron)", len(df_clean))
                    with col3:
                        st.metric("Registros Fallidos (No Cruzaron)", len(alerts))
                        
                    # Download Buttons
                    st.subheader("📥 Descargar Archivos Generados")
                    dl_col1, dl_col2 = st.columns(2)
                    with dl_col1:
                        st.download_button(
                            label="Descargar BASE PROLIJA.xlsx",
                            data=prolija_data,
                            file_name="BASE PROLIJA.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    with dl_col2:
                        st.download_button(
                            label="Descargar subir a predictivo.xlsx",
                            data=pred_data,
                            file_name="subir a predictivo.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                    # Display Alerts (failed rows)
                    st.subheader("⚠️ Reporte de Registros que NO Cruzaron")
                    if alerts:
                        df_alerts = pd.DataFrame(alerts)
                        st.dataframe(df_alerts, use_container_width=True)
                    else:
                        st.info("¡Excelente! Todas las filas cruzaron y se limpiaron correctamente.")
                        
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
else:
    st.info("💡 Por favor, sube un archivo Excel para comenzar.")
