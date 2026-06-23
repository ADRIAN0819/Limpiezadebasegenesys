import pandas as pd
import os

def completar_base():
    # Rutas de los archivos
    ruta_pend = os.path.join('PREDICTIVO', 'BASE PEND LLENAR.xlsx')
    ruta_datos = os.path.join('PREDICTIVO', 'LLENADO DE DATOS.xlsx')
    ruta_salida = os.path.join('PREDICTIVO', 'BASE PEND_COMPLETADA.xlsx')

    print(f"Cargando {ruta_pend}...")
    df_pend = pd.read_excel(ruta_pend, dtype=str)

    print(f"Cargando {ruta_datos}...")
    df_datos = pd.read_excel(ruta_datos, dtype=str)

    print("Realizando cruce de datos...")
    
    # Separar filas VRM y no-VRM en df_pend
    df_pend['BASE FINAL[REGLA]'] = df_pend['BASE FINAL[REGLA]'].fillna('').astype(str).str.strip()
    df_pend_vrm = df_pend[df_pend['BASE FINAL[REGLA]'] == 'VRM'].copy()
    df_pend_non_vrm = df_pend[df_pend['BASE FINAL[REGLA]'] != 'VRM'].copy()

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

    # Reemplazamos los valores en las columnas destino
    # Usamos where() para reemplazar solo cuando encontramos el valor en df_datos
    df_resultado['BASE FINAL[BASE WF.CELULAR/TELEFONO]'] = df_resultado['TELEFONO'].combine_first(df_resultado['BASE FINAL[BASE WF.CELULAR/TELEFONO]'])
    df_resultado['BASE FINAL[BASE WF.DNI]'] = df_resultado['DNI'].combine_first(df_resultado['BASE FINAL[BASE WF.DNI]'])
    df_resultado['BASE FINAL[BASE WF.CUENTA]'] = df_resultado['CUENTA'].combine_first(df_resultado['BASE FINAL[BASE WF.CUENTA]'])
    df_resultado['BASE FINAL[BASE WF.TARJETA]'] = df_resultado['TARJETA'].combine_first(df_resultado['BASE FINAL[BASE WF.TARJETA]'])
    df_resultado['BASE FINAL[BASE WF.NOMBRE DEL CLIENTE]'] = df_resultado['CLIENTE'].combine_first(df_resultado['BASE FINAL[BASE WF.NOMBRE DEL CLIENTE]'])

    # Eliminamos las columnas auxiliares que trajimos del merge
    df_resultado.drop(columns=['TELEFONO', 'DNI', 'CUENTA', 'TARJETA', 'CLIENTE'], inplace=True)

    print(f"Guardando archivo en {ruta_salida}...")
    # Guardamos el resultado
    df_resultado.to_excel(ruta_salida, index=False)
    print("¡Proceso completado exitosamente!")

if __name__ == '__main__':
    completar_base()