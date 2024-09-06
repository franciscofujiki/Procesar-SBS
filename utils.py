from sqlalchemy import create_engine
import pandas as pd
import requests
from io import BytesIO
import re
import calendar
from datetime import datetime

# Diccionario para convertir el nombre del mes a número del mes
meses = {
    'Enero': 1, 'Febrero': 2, 'Marzo': 3, 'Abril': 4, 'Mayo': 5, 'Junio': 6,
    'Julio': 7, 'Agosto': 8, 'Setiembre': 9, 'Octubre': 10, 'Noviembre': 11, 'Diciembre': 12
}

codigo = {
    "Enero": "en", "Febrero": "fe", "Marzo": "ma", "Abril": "ab", "Mayo": "my", "Junio": "jn",
    "Julio": "jl", "Agosto": "ag", "Setiembre": "se", "Octubre": "oc", "Noviembre": "no", "Diciembre": "di"
}

# Función para obtener código de 2 digitos del mes
def obtener_codigo_mes(mes):
    return codigo[mes]

# Función para obtener el último día del mes
def obtener_ultimo_dia_mes(año, mes):
    # Convertir el nombre del año en número
    año = int(año)
    # Convertir el nombre del mes al número del mes
    numero_mes = meses[mes]
    # Obtener el último día del mes
    ultimo_dia = calendar.monthrange(año, numero_mes)[1]
    # Formatear la fecha en formato dd-mm-yyyy
    return datetime(año, numero_mes, ultimo_dia).strftime('%Y-%m-%d')

def grabar_df_sql(server, database, df, tabla, esquema):
    try:
        # Crear la URL de conexión
        connection_string = f'mssql+pyodbc://{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes'
            
        # Crear un motor de conexión usando SQLAlchemy
        engine = create_engine(connection_string)

        # Insertar el DataFrame en la tabla 'mi_tabla' en la base de datos
        df.to_sql(tabla, con=engine, schema=esquema, if_exists='append', index=False)

    except Exception as e:
        print(f"Ocurrió un error: {e}")
        return None

def obtener_url_reporte(reporte):
    reportes = {
        # Balance General y Estado de Ganancias y Pérdidas	
        1: 'https://intranet2.sbs.gob.pe/estadistica/financiera/{año}/{mes_largo}/C-1101-{mes_corto}{año}.XLS',
        # Créditos a Actividades Empresariales por Sector Económico	
        2: 'https://intranet2.sbs.gob.pe/estadistica/financiera/{año}/{mes_largo}/C-1216-{mes_corto}{año}.XLS',
        # Créditos Directos según Tipo de Crédito y Situación	
        3: 'https://intranet2.sbs.gob.pe/estadistica/financiera/{año}/{mes_largo}/C-1228-{mes_corto}{año}.XLS',
        # Indicadores Financieros	
        4: 'https://intranet2.sbs.gob.pe/estadistica/financiera/{año}/{mes_largo}/C-1301-{mes_corto}{año}.XLS',
        # Créditos Directos según Situación	
        5: 'https://intranet2.sbs.gob.pe/estadistica/financiera/{año}/{mes_largo}/C-1207-{mes_corto}{año}.XLS',
        # Estructura del Activo	
        6: 'https://intranet2.sbs.gob.pe/estadistica/financiera/{año}/{mes_largo}/C-1217-{mes_corto}{año}.XLS',
        # Créditos Directos y Depósitos por Oficinas	
        7: 'https://intranet2.sbs.gob.pe/estadistica/financiera/{año}/{mes_largo}/C-1234-{mes_corto}{año}.XLS',
        # Número de Personal	
        8: 'https://intranet2.sbs.gob.pe/estadistica/financiera/{año}/{mes_largo}/C-1202-{mes_corto}{año}.XLS'
    }

    try:

        # Devolver el DataFrame
        return reportes[reporte]
    
    except Exception as e:
        print(f"Ocurrió un error: {e}")
        return None

def obtener_df_url(url, hoja=""):
    # Descargar el archivo Excel
    response = requests.get(url)

    # Verificar si la solicitud fue exitosa
    if response.status_code == 200:
        try:
            if hoja=="":
                # Leer el archivo Excel en un DataFrame
                df = pd.read_excel(BytesIO(response.content), engine='xlrd')
            else:
                df = pd.read_excel(BytesIO(response.content), sheet_name=hoja, engine='xlrd')

            # print(f"Archivo cargado exitosamente.")
            return df
        except Exception as e:
            print(f"Error al leer el archivo: {e}")
            return None
    else:
        print(f"Error al descargar el archivo. Código de estado:", response.status_code)
        return None

def limpiar_filas_columnas_vacias_df(df):
    # Eliminar filas donde todos los valores son NaN
    df = df.dropna(axis=0, how='all')
    # Eliminar columnas donde todos los valores son NaN
    df = df.dropna(axis=1, how='all')
    return df

def eliminar_filas_df_by_list(df, nombre_columna, lista):
    # Escapar los caracteres especiales en las subcadenas para evitar problemas de regex
    lista = [re.escape(subcadena) for subcadena in lista]

    # Usar una expresión regular para buscar cualquier subcadena en la lista
    patron = '|'.join(lista)

    # Convertir la columna 'Unnamed: 0' a string y aplicar el filtro
    df[nombre_columna] = df[nombre_columna].astype(str)

    # Eliminar las filas donde la columna 'A' contiene cualquiera de las subcadenas
    df = df[~df[nombre_columna].str.contains(patron, na=False)]

    return df

def eliminar_columnas_df_by_value(df, valor_a_eliminar, columna_inicio_index=0):

    # Filtrar columnas a partir de 'B' (inclusive)
    columnas_filtradas = df.columns[columna_inicio_index:]

    # Eliminar columnas a partir de 'B' si contienen el valor
    df_sin_columnas = df.drop(columns=[col for col in columnas_filtradas if df[col].astype(str).str.contains(valor_a_eliminar).any()])

    return df_sin_columnas

def establecer_fila_como_nombre_columnas_df(df, index_fila = 0):
    # Aplicar str.strip(), para limpiar espacios en blanco a los valores de la fila
    df.iloc[index_fila] = df.iloc[index_fila].apply(str.strip)

    # Tomar la primera fila como nombres de las columnas
    df.columns = df.iloc[index_fila]  # Asignar la primera fila como nombres de columnas

    # Eliminar la primera fila
    df = df.drop(df.index[index_fila]).reset_index(drop=True)  # Restablecer los índices

    return df