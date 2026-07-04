import pandas as pd
import os
import numpy as np

def generar_base_cruce():
    print("Cargando archivos...")
    
    # Rutas dinámicas
    predictivo_dir = 'PREDICTIVO'
    base_input_dir = os.path.join(predictivo_dir, 'BASE')
    
    archivos_lista = [f for f in os.listdir(base_input_dir) if 'lista' in f.lower() and f.endswith('.csv')]
    if not archivos_lista:
        print("Error: No se encontró ningún archivo de lista de alertas.")
        return
    ruta_lista = os.path.join(base_input_dir, archivos_lista[0])
    
    archivos_interacc = [f for f in os.listdir(base_input_dir) if 'interacciones' in f.lower() and f.endswith('.csv')]
    if not archivos_interacc:
        print("Error: No se encontró ningún archivo de interacciones.")
        return
    ruta_interacc = os.path.join(base_input_dir, archivos_interacc[0])

    resultados_dir = os.path.join(predictivo_dir, 'RESULTADOS')
    os.makedirs(resultados_dir, exist_ok=True)
    ruta_salida = os.path.join(resultados_dir, 'BASE CRUCE PRED_NUEVA.xlsx')

    print(f"Usando lista: {archivos_lista[0]}")
    print(f"Usando interacciones: {archivos_interacc[0]}")

    # ===== ORQUESTACIÓN BASE PEND =====
    while True:
        tipo_base_pend = input("¿Vas a usar 'BASE PEND LLENAR' (con cruce de datos) o 'BASE PEND LISTA'? (Escribe 1 para LLENAR, 2 para LISTA): ").strip()
        if tipo_base_pend in ["1", "2"]:
            break
        print("Opción no válida. Escribe 1 o 2.")
        
    if tipo_base_pend == "1":
        ruta_pend_llenar = os.path.join(base_input_dir, 'BASE PEND LLENAR.xlsx')
        ruta_datos = os.path.join(base_input_dir, 'LLENADO DE DATOS.xlsx')
        print(f"Completando BASE PEND... Cargando {ruta_pend_llenar} y {ruta_datos}")
        df_pend_temp = pd.read_excel(ruta_pend_llenar, dtype=str)
        df_datos = pd.read_excel(ruta_datos, dtype=str)
        
        # Separar filas VRM y no-VRM en df_pend_temp
        df_pend_temp['BASE FINAL[REGLA]'] = df_pend_temp['BASE FINAL[REGLA]'].fillna('').astype(str).str.strip()
        df_pend_vrm = df_pend_temp[df_pend_temp['BASE FINAL[REGLA]'] == 'VRM'].copy()
        df_pend_non_vrm = df_pend_temp[df_pend_temp['BASE FINAL[REGLA]'] != 'VRM'].copy()
        
        # 1. Procesamiento para VRM (Cruce por TARJETA)
        df_pend_vrm['BASE FINAL[TARJETA]'] = df_pend_vrm['BASE FINAL[TARJETA]'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df_datos_vrm = df_datos.copy()
        df_datos_vrm['BASE FINAL[TARJETA]'] = df_datos_vrm['BASE FINAL[TARJETA]'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df_datos_unica_vrm = df_datos_vrm.drop_duplicates(subset=['BASE FINAL[TARJETA]'], keep='first')
        columnas_vrm = ['BASE FINAL[TARJETA]', 'TELEFONO', 'DNI', 'CUENTA', 'TARJETA', 'CLIENTE']
        df_datos_unica_vrm = df_datos_unica_vrm[columnas_vrm]
        df_res_vrm = pd.merge(df_pend_vrm, df_datos_unica_vrm, on='BASE FINAL[TARJETA]', how='left', suffixes=('', '_datos'))
        
        # 2. Procesamiento para no-VRM (Cruce por ID)
        df_pend_non_vrm['BASE FINAL[ID]'] = df_pend_non_vrm['BASE FINAL[ID]'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df_datos_non_vrm = df_datos.copy()
        df_datos_non_vrm['BASE FINAL[ID]'] = df_datos_non_vrm['BASE FINAL[ID]'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df_datos_unica_non_vrm = df_datos_non_vrm.drop_duplicates(subset=['BASE FINAL[ID]'], keep='first')
        columnas_non_vrm = ['BASE FINAL[ID]', 'TELEFONO', 'DNI', 'CUENTA', 'TARJETA', 'CLIENTE']
        df_datos_unica_non_vrm = df_datos_unica_non_vrm[columnas_non_vrm]
        df_res_non_vrm = pd.merge(df_pend_non_vrm, df_datos_unica_non_vrm, on='BASE FINAL[ID]', how='left', suffixes=('', '_datos'))
        
        # Concatenar resultados
        df_resultado = pd.concat([df_res_vrm, df_res_non_vrm], ignore_index=True)
        
        df_resultado['BASE FINAL[BASE WF.CELULAR/TELEFONO]'] = df_resultado['TELEFONO'].combine_first(df_resultado['BASE FINAL[BASE WF.CELULAR/TELEFONO]'])
        df_resultado['BASE FINAL[BASE WF.DNI]'] = df_resultado['DNI'].combine_first(df_resultado['BASE FINAL[BASE WF.DNI]'])
        df_resultado['BASE FINAL[BASE WF.CUENTA]'] = df_resultado['CUENTA'].combine_first(df_resultado['BASE FINAL[BASE WF.CUENTA]'])
        df_resultado['BASE FINAL[BASE WF.TARJETA]'] = df_resultado['TARJETA'].combine_first(df_resultado['BASE FINAL[BASE WF.TARJETA]'])
        df_resultado['BASE FINAL[BASE WF.NOMBRE DEL CLIENTE]'] = df_resultado['CLIENTE'].combine_first(df_resultado['BASE FINAL[BASE WF.NOMBRE DEL CLIENTE]'])
        
        for col in ['BASE FINAL[BASE WF.CELULAR/TELEFONO]', 'BASE FINAL[BASE WF.CUENTA]', 'BASE FINAL[BASE WF.DNI]', 'BASE FINAL[BASE WF.TARJETA]']:
            df_resultado[col] = df_resultado[col].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '')
            
        df_resultado.drop(columns=['TELEFONO', 'DNI', 'CUENTA', 'TARJETA', 'CLIENTE'], inplace=True)
        df_pend = df_resultado
    else:
        ruta_pend_lista = os.path.join(base_input_dir, 'BASE PEND LISTA.xlsx')
        print(f"Cargando {ruta_pend_lista} directamente...")
        df_pend = pd.read_excel(ruta_pend_lista, dtype=str)

    # ===== FIN ORQUESTACIÓN =====

    import csv
    # Cargar datos
    df_lista = pd.read_csv(ruta_lista, dtype=str)
    
    # Fix provisorio por si el archivo de lista viene en una sola columna encerrado en comillas dobles
    if len(df_lista.columns) == 1 and 'TELEFONO' in df_lista.columns[0]:
        df_lista = pd.read_csv(ruta_lista, dtype=str, quoting=csv.QUOTE_NONE, sep=',')
        df_lista.columns = [c.replace('"', '') for c in df_lista.columns]
        for col in df_lista.columns:
            df_lista[col] = df_lista[col].astype(str).str.replace('"', '', regex=False)
        print("MUESTRA LIMPIA:", df_lista['TELEFONO'].iloc[0])
                
    try:
        df_interacc = pd.read_csv(ruta_interacc, sep=';', dtype=str)
        if len(df_interacc.columns) <= 1:
            raise ValueError("Separador incorrecto")
    except Exception:
        df_interacc = pd.read_csv(ruta_interacc, sep=',', dtype=str)

    print("Procesando llaves...")
    # Limpiar na de las listas
    df_lista = df_lista.fillna("")
    
    print("Cruzando con BASE PEND (Paso Temporal para obtener HORATRX)...")
    # Para formar el mismo nivel de "llave PEND" en el archivo de lista, debemos obtener la hora_trx de la BASE PEND
    # Hacemos un cruce preliminar solo por tarjeta/cuenta/dni para traer la hora.
    # Como la llave original era sin hora, usaremos temporalmente esa para traer la hora.
    df_lista['llave PEND_temp'] = df_lista['TELEFONO'] + df_lista['CUENTA'] + df_lista['DNI'] + df_lista['TARJETA']
    
    df_pend_temp_key = df_pend.copy()
    df_pend_temp_key['llave PEND_temp'] = df_pend_temp_key['BASE FINAL[BASE WF.CELULAR/TELEFONO]'] + df_pend_temp_key['BASE FINAL[BASE WF.CUENTA]'] + df_pend_temp_key['BASE FINAL[BASE WF.DNI]'] + df_pend_temp_key['BASE FINAL[BASE WF.TARJETA]']
    # Se incluye la hora para no perder las transacciones de diferentes horas del mismo cliente
    df_pend_temp_key['hora_limpia_temp'] = df_pend_temp_key['BASE FINAL[HORATRX]'].astype(str).str.extract(r'(\d{2}:\d{2})')[0].fillna("")
    df_pend_temp_key = df_pend_temp_key.drop_duplicates(subset=['llave PEND_temp', 'hora_limpia_temp'], keep='first')
        
    df_lista = pd.merge(df_lista, df_pend_temp_key[['llave PEND_temp', 'BASE FINAL[HORATRX]']], on='llave PEND_temp', how='left')
    hora_transaccion_lista = df_lista['BASE FINAL[HORATRX]'].astype(str).str.extract(r'(\d{2}:\d{2})')[0].fillna("")
    
    # IMPORTANTE: Descartar esta columna temporal para no crear conflicto en el merge final con df_pend
    df_lista = df_lista.drop(columns=['BASE FINAL[HORATRX]'])

    # 1. Crear llave PEND completa y definitiva en df_lista
    df_lista['llave PEND'] = df_lista['llave PEND_temp'] + hora_transaccion_lista

    # 2. FECHA Y HORA DE LISTA PRED (de y2 que es CallRecordLastAttempt-TELEFONO)
    # Ejemplo Y2: "2026-06-09T02:32:04.615Z"
    # IMPORTANTE: Replicamos formula de Excel EXTRAE(Y2;ENCONTRAR("T";Y2)+1;5) que trunca a nivel de minutos
    call_record_str = df_lista['CallRecordLastAttempt-TELEFONO'].astype(str)
    call_record_trunc = call_record_str.str[:16] # Toma hasta "YYYY-MM-DDTHH:MM"
    call_record_dates = pd.to_datetime(call_record_trunc, format='%Y-%m-%dT%H:%M', errors='coerce')
    # Restar 5 horas
    call_record_dates_lima = call_record_dates - pd.Timedelta(hours=5)
    df_lista['FECHA Y HORA DE LISTA PRED'] = call_record_dates_lima.dt.strftime('%d/%m/%Y %H:%M:%S').fillna("")

    # 3. LLAVE GENESYS en df_lista (=O2&J2 -> inin-outbound-id & FECHA Y HORA DE LISTA PRED)
    df_lista['LLAVE GENESYS'] = df_lista['inin-outbound-id'] + df_lista['FECHA Y HORA DE LISTA PRED']

    # 4. Crear llave en df_interacciones (Identificación de contacto & Fecha de finalización formateada)
    dt_fin_interacc = pd.to_datetime(df_interacc['Fecha de finalización'], format='%d/%m/%y %H:%M', errors='coerce')
    dt_fin_interacc_str = dt_fin_interacc.dt.strftime('%d/%m/%Y %H:%M:%S').fillna("")
    df_interacc['LLAVE GENESYS'] = df_interacc['Identificación de contacto'].fillna("") + dt_fin_interacc_str
    
    # Quitar duplicados para asegurar mejor cruce (tomamos el primero en caso de duplicados)
    df_interacc_unica = df_interacc.drop_duplicates(subset=['LLAVE GENESYS'], keep='first')

    print("Cruzando con Interacciones...")
    # Cruzar para traer Usuarios - Alertados, Fecha de finalización
    df_lista = pd.merge(df_lista, df_interacc_unica[['LLAVE GENESYS', 'Usuarios - Alertados', 'Fecha de finalización']], 
                        on='LLAVE GENESYS', how='left')

    # 5. AG
    df_lista['AG'] = df_lista['Usuarios - Alertados'].fillna("ZLAO")
    df_lista['AG'] = df_lista['AG'].replace("", "ZLAO")

    # 6. FECHA - REGISTRO / HORA - REGISTRO
    # Extraemos del resultado y si es NA usamos la fecha origen del contacto (FECHA Y HORA DE LISTA PRED)
    dt_reg_inter = pd.to_datetime(df_lista['Fecha de finalización'], format='%d/%m/%y %H:%M', errors='coerce')
    
    # Manejar los fallos en fechas usando la fecha originada de genesys
    fecha_genesys = pd.to_datetime(df_lista['FECHA Y HORA DE LISTA PRED'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
    
    dt_reg_final = dt_reg_inter.combine_first(fecha_genesys)
    
    df_lista['FECHA  - REGISTRO'] = dt_reg_final.dt.strftime('%d/%m/%Y').fillna("")
    df_lista['HORA - REGISTRO'] = dt_reg_final.dt.strftime('%H:%M:%S').fillna("")

    print("Cruzando con BASE PEND...")
    # Crear la llave en BASE PEND incluyendo hora trx para evitar sobreescritura de duplicados de diferente tramo
    df_pend = df_pend.fillna("")
    
    # Intentamos primero extraer hora trx limpiándola directamente por si trae minutos extras o letras (ya que lo uniremos a la llave)
    hora_transaccion_limpia = df_pend['BASE FINAL[HORATRX]'].astype(str).str.extract(r'(\d{2}:\d{2})')[0].fillna("")
    df_pend['llave PEND'] = df_pend['BASE FINAL[BASE WF.CELULAR/TELEFONO]'] + df_pend['BASE FINAL[BASE WF.CUENTA]'] + df_pend['BASE FINAL[BASE WF.DNI]'] + df_pend['BASE FINAL[BASE WF.TARJETA]'] + hora_transaccion_limpia
    
    df_pend_unica = df_pend.drop_duplicates(subset=['llave PEND'], keep='first')
    
    # Cruzamos
    df_lista = pd.merge(df_lista, df_pend_unica[['llave PEND', 'BASE FINAL[FECHATRX]', 'BASE FINAL[HORATRX]', 'BASE FINAL[REGLA]', 'BASE FINAL[NOMBRECOMERCIO]']], 
                        on='llave PEND', how='left')
    
    # 7. fecha trx, hora trx
    # Base pend FECHATRX ya esta en formato excel, extraerlo
    # Manejador basico para hora trx que en el output sale como hh:mm
    df_lista['fecha trx'] = df_lista['BASE FINAL[FECHATRX]'].fillna("")
    
    # Extraer estrictamente el formato HH:MM (horas y minutos) ignorando segundos u otros caracteres
    df_lista['hora trx'] = df_lista['BASE FINAL[HORATRX]'].astype(str).str.extract(r'(\d{2}:\d{2})')[0]
    df_lista['hora trx'] = df_lista['hora trx'].fillna("")

    print("Cruzando variables adicionales...")
    # 8. Canal
    mapa_canal = {
        'FRM': 'FRM - FRM',
        'IZIPAY': 'MC - MasterCard (TC)',
        'SQL_BATCH_BLBM_0054': 'BM - Banca Móvil (App)',
        'SQL_NRT_BM_0052': 'BM - Banca Móvil (App)',
        'SQL_NRT_HHA_0047': 'INT-HBK - Multired Virtual (HBK)',
        'SQL_NRT_ATMH_0072': "ATM - ATM's Multired",
        'VRM': 'VRM - VISA (TMGDV)',
        'SQL_BATCH_HCS_0057': 'INT-HBK - Multired Virtual (HBK)',
        'SQL_BATCH_BLMP_0058': 'BATCH - MICROPAGOS - BATCH - MICROPAGOS',
        'SQL_BATCH_VN7_0061': 'BATCH VISA N7 - BATCH VISA N7',
        'SQL-TA-ATM-0002': "ATM - ATM's Multired",
        'SQL_BATCH_GIRO_0069': 'BATCH \u2013 GIROS - BATCH \u2013 GIROS',
        'SQL_BATCH_MTCV_0059': 'MULTICANAL - MULTICANAL'
    }
    df_lista['canal'] = df_lista['BASE FINAL[REGLA]'].map(mapa_canal).fillna("")

    # 9. DNI8 (rellenar ceros)
    df_lista['DNI8'] = df_lista['DNI'].str.zfill(8)

    # Pedir al usuario por consola si es PENDIENTES o CIERRE
    while True:
        tipo_calif = input("¿Es PENDIENTES o CIERRE?: ").strip().upper()
        if tipo_calif in ["PENDIENTES", "CIERRE"]:
            break
        print("Opción no válida. Por favor escribe PENDIENTES o CIERRE.")

    valor_comunicacion = input("Ingresa el valor para COMUNICACION 1: ").strip()
    nombre_archivo_final = input("Ingresa el nombre para el archivo final (sin extensión): ").strip()
    
    if not nombre_archivo_final:
        nombre_archivo_final = "MASIVO_PREDICTIVO"

    if tipo_calif == "PENDIENTES":
        df_lista['CALIFICACION'] = "PEND - Alerta pendiente(Temporal)"
    else:
        # Cargar lista de comercios riesgosos
        ruta_comercios = 'COMERCIOS RIESGOSOS.xlsx'
        if os.path.exists(ruta_comercios):
            df_comercios = pd.read_excel(ruta_comercios, dtype=str)
            comercios_riesgosos = df_comercios['NOMBRE'].fillna("").astype(str).str.strip().str.upper().tolist()
        else:
            print(f"Advertencia: No se encontró {ruta_comercios}. Se asignará NFCNC a todos.")
            comercios_riesgosos = []
            
        comercio_pend = df_lista['BASE FINAL[NOMBRECOMERCIO]'].fillna("").astype(str).str.strip().str.upper()
        es_riesgoso = comercio_pend.isin(comercios_riesgosos)
        
        df_lista['CALIFICACION'] = np.where(es_riesgoso, 
                                           "FAG - Fraude por análisis del gestor", 
                                           "NFCNC - No fraude cliente no contesta(Definitivo)")

    df_lista['1 COMUNICACIÓN'] = "LLAMADA PREDICTIVA"
    df_lista['origen'] = "MANUAL"

    # 11. nivel de respuesta
    mapa_resp = {
        'ININ-OUTBOUND-BUSY': 'Cliente no contesta',
        'ININ-OUTBOUND-DISCONNECT': 'Cliente no contesta',
        'ININ-OUTBOUND-FAILED-TO-REACH-AGENT': 'Telefono Inoperativo - No recibe llamadas',
        'ININ-OUTBOUND-FAILED-TO-REACH-FLOW': 'Telefono Inoperativo - No recibe llamadas',
        'ININ-OUTBOUND-FAX': 'Cliente no contesta',
        'ININ-OUTBOUND-INVALID-PHONE-NUMBER': 'Telefono no existe',
        'ININ-OUTBOUND-NO-ANSWER': 'Cliente no contesta',
        'ININ-OUTBOUND-SIT-CALLABLE': 'Cliente no contesta',
        'ININ-OUTBOUND-STUCK-INTERACTION': 'Cliente no contesta',
        'Llamada con Mensaje': 'Buzon de voz',
        'Default Wrap-up Code': 'Cliente no contesta',
        'ININ-OUTBOUND-AMBIGUOUS': 'Cliente no contesta',
        'ININ-OUTBOUND-INTERNAL-ERROR-SKIPPED': 'Cliente no contesta',
        'ININ-OUTBOUND-SIT-UNCALLABLE': 'Cliente no contesta',
        'ININ-OUTBOUND-MACHINE': 'Cliente no contesta',
        'ININ-WRAP-UP-TIMEOUT': 'Cliente no contesta'
    }
    df_lista['nivel de respuesta'] = df_lista['CallRecordLastResult-TELEFONO'].map(mapa_resp).fillna("")

    # Formar las columnas en el orden solicitado
    columnas_finales = [
        "llave PEND", "LLAVE GENESYS", "AG", "FECHA  - REGISTRO", "HORA - REGISTRO", 
        "fecha trx", "hora trx", "canal", "DNI8", "FECHA Y HORA DE LISTA PRED", 
        "CALIFICACION", "1 COMUNICACIÓN", "nivel de respuesta", "origen", 
        "inin-outbound-id", "TELEFONO", "CUENTA", "DNI", "TARJETA", "NOMBRE", 
        "ContactCallable", "ContactableByVoice", "ContactableBySms", "ContactableByEmail", 
        "CallRecordLastAttempt-TELEFONO", "CallRecordLastResult-TELEFONO", "CallRecordLastAgentWrapup-TELEFONO", 
        "SmsLastAttempt-TELEFONO", "SmsLastResult-TELEFONO", "Callable-TELEFONO", 
        "ContactableByVoice-TELEFONO", "ContactableBySms-TELEFONO", "ContactableByWhatsApp", 
        "CallRecordLastAttemptCampaign-TELEFONO"
    ]
    
    # SmsLastAttemptCampaign-TELEFONO esta en el dataset base, podemos añadirlo al final
    if 'SmsLastAttemptCampaign-TELEFONO' in df_lista.columns:
        columnas_finales.append('SmsLastAttemptCampaign-TELEFONO')

    # Seleccionar solo las columnas ordenadas
    df_final = df_lista[columnas_finales]

    print(f"Guardando archivo en {ruta_salida}...")
    df_final.to_excel(ruta_salida, index=False)
    
    print("Generando archivo final de gestión...")
    df_masivo = pd.DataFrame()
    df_masivo['EXPEDIENTE'] = [""] * len(df_final)
    df_masivo['FECHA'] = df_final['FECHA  - REGISTRO']
    df_masivo['HORA'] = df_final['HORA - REGISTRO']
    df_masivo['TIPO DE GESTIÓN'] = "Outbound"
    df_masivo['GESTIÓN'] = "Call Out Monitoreo"
    df_masivo['GESTIÓN MOTIVO'] = ""
    df_masivo['FECHA Y HORA DE ALERTA'] = ""
    df_masivo['FECHA Y HORA DE ATENCION'] = ""
    df_masivo['Of. Banco de la Nación'] = ""
    df_masivo['Anexo Interno'] = ""
    df_masivo['Nombre Funcionario BN'] = ""
    df_masivo['CELULAR/TELEFONO'] = df_final['TELEFONO']
    df_masivo['DNI'] = df_final['DNI']
    df_masivo['NOMBRE DEL CLIENTE'] = df_final['NOMBRE']
    df_masivo['CORREO'] = ""
    df_masivo['CUENTA EMISORA'] = df_final['CUENTA']
    df_masivo['CUENTA RECEPTORA'] = ""
    df_masivo['CUENTA RECEPTORA 2'] = ""
    df_masivo['CUENTA RECEPTORA 3'] = ""
    df_masivo['NRO. GIRO 1'] = ""
    df_masivo['NRO. GIRO 2'] = ""
    df_masivo['NRO. GIRO 3'] = ""
    df_masivo['NRO. GIRO 4'] = ""
    df_masivo['NRO. GIRO 5'] = ""
    df_masivo['TARJETA'] = df_final['TARJETA']
    df_masivo['N° BLQ'] = ""
    df_masivo['FECHA(BLQ_VIG)'] = ""
    df_masivo['CANAL'] = df_final['canal']
    df_masivo['REGLA MONITOREO'] = ""
    df_masivo['REGLA O PARAMETRO DE BLOQUEO'] = ""
    df_masivo['SITUACION_BDUC'] = "ACTUALIZADO"
    df_masivo['CALIFICACION'] = df_final['CALIFICACION']
    df_masivo['OPCION CALIFICACION'] = ""
    df_masivo['FECHA Y HORA DE ENVIO DE CORREO'] = ""
    df_masivo['IMPORTE DE FRAUDE'] = ""
    
    mapa_gestor = {
        'ZLAO': 'DAVID JOSUE ZAMBRANO LEON',
        'AG0008 David Zambrano': 'DAVID JOSUE ZAMBRANO LEON',
        'AG0018 ELIANA MIRANDA': 'ELIANA ESTELA MIRANDA POMACONDOR',
        'AG0024 Antonia Montoya': 'ANTONIA MONTOYA HUAMANI',
        'AG0030 TANIA GUERRERO': 'TANIA LISSET GUERRERO ROMERO',
        'AG0108 MIGUEL NAVARRO': 'MIGUEL ANGEL NAVARRO SALAZAR',
        'AG0124 TERESA CABANILLAS': 'TERESA DEL CARMEN CABANILLAS JAVIER',
        'AG0126 JOSUE OLAYA': 'JOSUE MANUEL OLAYA DOMÍNGUEZ',
        'AG0129 JANICE MENDOZA': 'JANICE MARJURIE MENDOZA ZAVALA',
        'AG0130 RAQUEL MARIANO': 'RAQUEL AYME MARIANO ORIHUELA',
        'AG0134 MANUEL PEÑA': 'MANUEL MARTIN PEÑA BENITES',
        'AG0142 Sandra Serrano': 'SANDRA SERRANO MOLINA',
        'AG0151 YOSEP SANGAY': 'YOSEP SANGAY VEGA',
        'AG0152 VALERIA ESPINOZA': 'VALERIA BELEN ESPINOZA DIAZ',
        'AG0173 JHONATAN MENDOZA': 'JHONATAN ENRIQUE MENDOZA CHAVEZ',
        'AG0176 MAURICIO SALCEDO': 'MAURICIO EDSON SALCEDO ENRIQUEZ',
        'AG0197 Lizbeth Chavez': 'LIZBETH CHAVEZ TORRES',
        'AG0215 ROSARIO HUATUCO': 'ROSARIO DEL PILAR HUATUCO QUISPE',
        'AG0229 NEFTALI SUYURI': 'NEFTALI AARON SUYURI SUYURI',
        'AG0263 DERECK MINAYA': 'DERECK ENRIQUE MINAYA CHU'
    }
    
    def mapear_gestor(val):
        val = str(val).strip()
        # Buscar por substring exacto
        for k, v in mapa_gestor.items():
            if k in val:
                return v
        return '///'

    df_masivo['GESTOR'] = df_final['AG'].apply(mapear_gestor)

    df_masivo['COMENTARIO'] = ""
    df_masivo['DIA DE EVENTO'] = df_final['fecha trx']
    df_masivo['HORA DE EVENTO'] = df_final['hora trx']
    df_masivo['TIEMPO DE DURACION'] = "Activa"
    df_masivo['G2'] = ""
    df_masivo['FALLECIMIENTO (SI/NO)'] = "NO"
    df_masivo['APLICACIÓN'] = ""
    df_masivo['Columna2'] = ""
    df_masivo['Columna3'] = ""
    df_masivo['RESULTADO DE LLAMADA'] = "No Contactado"
    df_masivo['NIVEL DE RESPUESTA'] = df_final['nivel de respuesta']
    df_masivo['MOTIVO DE ATENCION'] = ""
    df_masivo['VALIDACION DE IDENTIDAD'] = ""
    df_masivo['TIPO DE TRANSACCION'] = ""
    df_masivo['IMPORTE RECUPERADO'] = ""
    df_masivo['NUMERO DE RECLAMO'] = ""
    df_masivo['TIPO DE FRAUDE'] = ""
    df_masivo['VIGILANCIA DE CUENTA'] = ""
    df_masivo['LEVANTAMIENTO VIGILANCIA'] = "NO APLICA"
    df_masivo['SOLUCION DE CASO'] = "Solucionada"
    df_masivo['FECHA Y HORA BLOQUEO'] = ""
    df_masivo['FECHA Y HORA DESBLOQUEO'] = ""
    df_masivo['COMUNICACION 1'] = valor_comunicacion
    df_masivo['COMUNICACION 2'] = ""
    df_masivo['COMUNICACION 3'] = ""
    df_masivo['FECHA MODIFICACION'] = ""
    df_masivo['MATERIALIZACION DE FRAUDE'] = ""
    df_masivo['SALDO DISPONIBLE'] = ""
    df_masivo['TIPO DE CUENTA'] = ""
    df_masivo['CONCLUSION'] = ""
    df_masivo['FECHA BLOQUEO DE TARJETA'] = ""

    # Ubicar cuáles filas de df_final tienen campos vacíos cruciales (que originalmente cruzaron como N/D pero fueron formateadas a vacío)
    # Ejemplo: canal vacío, nivel de respuesta vacío, o la hora trx está vacía
    mascara_invalidos = (df_final['canal'] == "") | (df_final['nivel de respuesta'] == "") | (df_final['hora trx'] == "") | (df_final['fecha trx'] == "") | (df_masivo['CANAL'] == "")

    # Limpiar cualquier "#N/D" sobrante a ""
    df_final = df_final.replace("#N/D", "")
    df_masivo = df_masivo.replace("#N/D", "")

    # Filtrar df_masivo quitando las filas que no lograron un cruce válido
    df_masivo = df_masivo[~mascara_invalidos]

    ruta_masivo = os.path.join(resultados_dir, f"{nombre_archivo_final}.xlsx")
    print(f"Guardando archivo final en {ruta_masivo}...")
    df_masivo.to_excel(ruta_masivo, index=False)

    print("¡Proceso completado exitosamente!")

if __name__ == '__main__':
    generar_base_cruce()
