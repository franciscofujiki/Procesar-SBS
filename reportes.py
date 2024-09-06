from utils import obtener_url_reporte, obtener_df_url, limpiar_filas_columnas_vacias_df, \
                    eliminar_filas_df_by_list, eliminar_columnas_df_by_value, \
                    establecer_fila_como_nombre_columnas_df, grabar_df_sql, obtener_ultimo_dia_mes, \
                    obtener_codigo_mes

import numpy as np
import pandas as pd

# Balance General: notebook
def generarReporteBalanceGeneral(anio, mes, reporte=1):
    # Último deía del mes
    ult_dia_mes = obtener_ultimo_dia_mes(anio, mes)

    # Codigo de 2 dígitos
    codigo = obtener_codigo_mes(mes)

    # Obtenemos URL
    url = obtener_url_reporte(reporte)

    url = url.replace("{mes_largo}",mes).replace("{mes_corto}",codigo).replace("{año}",anio)

    # Obtenemos Excel
    df = obtener_df_url(url, hoja="bg_cm")
    df = limpiar_filas_columnas_vacias_df(df)
    df.reset_index(inplace=True, drop=True)

    # Lista de subcadenas a buscar para eliminar las filas
    subcadenas_a_eliminar = ['Balance General por', '(En Miles de Soles)', '00:00:00', 'Tipo de Cambio',
                            'A partir de ', 'Actualizado al', 'Actualizado el', '(En  Miles de Soles)', 
                            'Incluye gastos devengados por pagar.','Con relación a la CMAC']

    df = eliminar_filas_df_by_list(df, 'Unnamed: 0', subcadenas_a_eliminar)
    df.reset_index(inplace=True, drop=True)

    # Eliminamos columnas repetidas donde se encuentre la palabra Activo
    columna_inicio_index = 1
    valor_a_eliminar = "Activo"

    df = eliminar_columnas_df_by_value(df, valor_a_eliminar, columna_inicio_index)
    df.reset_index(inplace=True, drop=True)

    # Seleccionar la fila que deseas corregir, en este caso, por ejemplo, la fila 0
    fila_a_corregir = 0

    # Aplicar el método 'ffill' para rellenar los valores nulos hacia adelante
    df.loc[fila_a_corregir] = df.loc[fila_a_corregir].ffill()

    # Reemplazar el valor 'nan' en la columna 'Unnamed: 0' con NaN
    df['Unnamed: 0'] = df['Unnamed: 0'].replace('nan', np.nan)

    # Concatenar las filas 0 y 1
    df.loc[0] = df.loc[0].fillna('').astype(str) + " " + df.loc[1].fillna('').astype(str)

    # Eliminar la segunda fila después de la concatenación
    df = df.drop(index=1)

    df.reset_index(inplace=True, drop=True)

    # Eliminar la ultima columna
    if df.iloc[0,-1].strip() == "TOTAL CAJAS MUNICIPALES":
        df = df.iloc[:, :-1]

    # Eliminamos donde haya NaN
    df = df.dropna(subset=['Unnamed: 0'])

    # Para crear la columna Grupo se usa np.where() para asignar los valores según las condiciones
    df['Grupo'] = np.where(df['Unnamed: 0'].str.contains('Activo'), 'Activo',
                    np.where(df['Unnamed: 0'].str.contains('Pasivo'), 'Pasivo', None))

    # Rellenar los valores nulos en la columna 'Grupo' hacia adelante
    df['Grupo'] = df['Grupo'].ffill()

    # Eliminamos filas que aún quedan con la Lista de subcadenas a buscar para eliminar las filas
    subcadenas_a_eliminar = ['Pasivo']

    df = eliminar_filas_df_by_list(df, 'Unnamed: 0', subcadenas_a_eliminar)
    df.reset_index(inplace=True, drop=True)

    # Colocamos el nombre de la 1er fila para la 1er columna y ultima
    df.iloc[0,0] = 'Item'
    df.iloc[0,-1] = 'Grupo'

    df = establecer_fila_como_nombre_columnas_df(df)

    df.reset_index(inplace=True)

    # Usar pd.melt() para convertir banco1 y banco2 en formato largo
    df = pd.melt(df, id_vars=['Grupo', 'index', 'Item'], value_vars=df.columns.difference(["Item", 'index', "Grupo"]),
                    var_name='Entidad', value_name='Valor')

    # Reemplazar el carácter '*' con un valor vacío en la columna 'col1'
    df['Item'] = df['Item'].str.replace('*', '', regex=False)

    df['Valor'] = pd.to_numeric(df['Valor'])

    df['AnioMes'] = ult_dia_mes

    # Mapeo entre los nombres de las columnas del DataFrame y los de la tabla en SQL
    mapeo_columnas = {
        'AnioMes': 'dFecha',
        'Grupo': 'cGrupoBalance',
        'index': 'nOrdenItemBalance',
        'Item': 'cItemBalance',
        'Entidad': 'cEntidad',
        'Valor': 'nValor'
    }

    # Renombrar las columnas del DataFrame para que coincidan con los nombres en SQL
    df = df.rename(columns=mapeo_columnas)

    return df

# Estado de Ganancias y Perdidas: notebook
def generarReporteEstadoGananciasPerdidas(anio, mes, reporte=1):
    # Último deía del mes
    ult_dia_mes = obtener_ultimo_dia_mes(anio, mes)
    
    # Codigo de 2 dígitos
    codigo = obtener_codigo_mes(mes)
    
    # Obtenemos URL
    url = obtener_url_reporte(reporte)

    url = url.replace("{mes_largo}",mes).replace("{mes_corto}",codigo).replace("{año}",anio)

    # Obtenemos Excel
    df = obtener_df_url(url, hoja="gyp_cm")

    # Si no encuentra dataframe
    if df is None:
        return None

    # Eliminamos filas y columnas vacias
    df = limpiar_filas_columnas_vacias_df(df)
    df.reset_index(inplace=True, drop=True)

    # Lista de subcadenas a buscar para eliminar las filas
    subcadenas_a_eliminar = ['Estado de Ganancias y Pérdidas', '(En Miles de Soles)', '00:00:00', 'Tipo de Cambio',
                            'A partir de ', 'Actualizado al', '(En  Miles de Soles)', 
                            'Incluye gastos devengados por pagar.','Con relación a la CMAC']

    df = eliminar_filas_df_by_list(df, 'Unnamed: 0', subcadenas_a_eliminar)
    df.reset_index(inplace=True, drop=True)

    # Si no contiene CMAC elimina la fila
    if pd.isna(df.iloc[0,1]):
        df = df.drop(df.index[0]).reset_index(drop=True)

    df = limpiar_filas_columnas_vacias_df(df)
    df.reset_index(inplace=True, drop=True)

    # Eliminamos columnas que se repiten
    columna_inicio_index = 1
    valor_a_eliminar = "INGRESOS FINANCIEROS"

    df = eliminar_columnas_df_by_value(df, valor_a_eliminar, columna_inicio_index)
    df.reset_index(inplace=True, drop=True)

    # Seleccionar la fila que deseas corregir, en este caso, por ejemplo, la fila 0
    fila_a_corregir = 0

    # Aplicar el método 'ffill' para rellenar los valores nulos hacia adelante
    df.loc[fila_a_corregir] = df.loc[fila_a_corregir].ffill()

    # Concatenar las filas 0 y 1
    df.loc[0] = df.loc[0].fillna('').astype(str) + " " + df.loc[1].fillna('').astype(str)

    df.iloc[0,0] = 'Item'

    # Eliminar la segunda fila después de la concatenación
    df = df.drop(index=1)

    df.reset_index(inplace=True, drop=True)

    # Tomar la primera fila como nombres de las columnas
    df = establecer_fila_como_nombre_columnas_df(df)

    df.reset_index(inplace=True)

    # Usar pd.melt() para convertir banco1 y banco2 en formato largo
    df = pd.melt(df, id_vars=['index', 'Item'], value_vars=df.columns.difference(["Item"]),
                    var_name='Entidad', value_name='Valor')

    df['Valor'] = pd.to_numeric(df['Valor'])

    df['AnioMes'] = ult_dia_mes

    # Mapeo entre los nombres de las columnas del DataFrame y los de la tabla en SQL
    mapeo_columnas = {
        'AnioMes': 'dFecha',
        'index': 'nOrdenItemGananciasPerdidas',
        'Item': 'cItemGananciasPerdidas',
        'Entidad': 'cEntidad',
        'Valor': 'nValor'
    }

    # Renombrar las columnas del DataFrame para que coincidan con los nombres en SQL
    df = df.rename(columns=mapeo_columnas)

    return df

# Créditos a Actividades Empresariales por Sector Económico: notebook2
def generarReporteSectorEconomicoEmpresariales(anio, mes, reporte=2):
    # Último deía del mes
    ult_dia_mes = obtener_ultimo_dia_mes(anio, mes)

    # Codigo de 2 dígitos
    codigo = obtener_codigo_mes(mes)

    # Obtenemos URL del reporte
    url = obtener_url_reporte(reporte)

    url = url.replace("{mes_largo}",mes).replace("{mes_corto}",codigo).replace("{año}",anio)

    # Obtenemos Excel
    df = obtener_df_url(url)

    # Si no encuentra dataframe
    if df is None:
        return None

    df = limpiar_filas_columnas_vacias_df(df)
    df.reset_index(inplace=True, drop=True)

    # Resetear los nombres de las columnas a 'col_1', 'col_2', 'col_3'
    df.columns = [f'col_{i+1}' for i in range(df.shape[1])]

    # Lista de subcadenas a buscar para eliminar las filas
    subcadenas_a_eliminar = ['Nota: ', '00:00:00', '(En miles de soles)', 'http://intranet1',
                             'Actualizado al', '(En miles de nuevos soles)', 'Actualizado el']

    df = eliminar_filas_df_by_list(df, 'col_1', subcadenas_a_eliminar)
    df.reset_index(inplace=True, drop=True)

    # Colocamos el nombre de la 1er fila para la 1er columna y ultima
    df = establecer_fila_como_nombre_columnas_df(df)

    df.reset_index(inplace=True)

    # Usar pd.melt() para convertir banco1 y banco2 en formato largo
    df = pd.melt(df, id_vars=['index', 'Sector Económico'], value_vars=df.columns.difference(['index', 'Sector Económico']),
                    var_name='Entidad', value_name='Valor')

    df['Valor'] = pd.to_numeric(df['Valor'])

    df['AnioMes'] = ult_dia_mes

    # Mapeo entre los nombres de las columnas del DataFrame y los de la tabla en SQL
    mapeo_columnas = {
        'AnioMes': 'dFecha',
        'index': 'nOrdenItem',
        'Sector Económico': 'cSectorEconomico',
        'Entidad': 'cEntidad',
        'Valor': 'nValor'
    }

    # Renombrar las columnas del DataFrame para que coincidan con los nombres en SQL
    df = df.rename(columns=mapeo_columnas)

    return df

# Créditos Directos según Tipo de Crédito y Situación: notebook3
def generarReporteTipoCreditoSituacion(anio, mes, reporte=3):
    # Último deía del mes
    ult_dia_mes = obtener_ultimo_dia_mes(anio, mes)

    # Codigo de 2 dígitos
    codigo = obtener_codigo_mes(mes)

    # Obtenemos URL del reporte
    url = obtener_url_reporte(reporte)
    url = url.replace("{mes_largo}",mes).replace("{mes_corto}",codigo).replace("{año}",anio)

    # Obtenemos Excel
    df = obtener_df_url(url)

    # Si no encuentra dataframe
    if df == None:
        return None

    df = limpiar_filas_columnas_vacias_df(df)
    df.reset_index(inplace=True, drop=True)

    # Resetear los nombres de las columnas a 'col_1', 'col_2', 'col_3'
    df.columns = [f'col_{i+1}' for i in range(df.shape[1])]

    # Lista de subcadenas a buscar para eliminar las filas
    subcadenas_a_eliminar = ['Nota: ', '00:00:00', '(En miles de soles)', 'http://intranet1',
                            'Actualizado al', '(En miles de nuevos soles)', 'Actualizado el',
                            'Las definiciones de los tipos','A partir de','En el marco de la Emergencia',
                            'de la declaratoria de','presentaban más de','Actualizado']

    df = eliminar_filas_df_by_list(df, 'col_1', subcadenas_a_eliminar)
    df.reset_index(inplace=True, drop=True)

    # Reemplazar el valor 'nan' en la columna 'col2' con NaN
    df['col_1'] = df['col_1'].replace('nan', np.nan)

    # Rellenar los valores nulos en la columna 'col_1' hacia adelante
    df['col_1'] = df['col_1'].ffill()

    # Colocamos el nombre de la 1er fila para la 1er columna y ultima
    df = establecer_fila_como_nombre_columnas_df(df)

    df.reset_index(inplace=True)

    # Usar pd.melt() para convertir banco1 y banco2 en formato largo
    df = pd.melt(df, id_vars=['index', 'Tipo de crédito', 'Situación'], value_vars=df.columns.difference(['index', 'Tipo de crédito', 'Situación']),
                    var_name='Entidad', value_name='Valor')

    df['Valor'] = pd.to_numeric(df['Valor'])

    df['AnioMes'] = ult_dia_mes
    
    # Mapeo entre los nombres de las columnas del DataFrame y los de la tabla en SQL
    mapeo_columnas = {
        'AnioMes': 'dFecha',
        'index': 'nOrdenItem',
        'Tipo de crédito': 'cTipoCredito',
        'Situación': 'cSituacion',
        'Entidad': 'cEntidad',
        'Valor': 'nValor'
    }

    # Renombrar las columnas del DataFrame para que coincidan con los nombres en SQL
    df = df.rename(columns=mapeo_columnas)

    return df

# Indicadores Financieros: notebook4
def generarReporteIndicadoresFinancieros(anio, mes, reporte=4):
    # Último deía del mes
    ult_dia_mes = obtener_ultimo_dia_mes(anio, mes)

    # Codigo de 2 dígitos
    codigo = obtener_codigo_mes(mes)

    # Obtenemos URL del reporte
    url = obtener_url_reporte(reporte)
    url = url.replace("{mes_largo}",mes).replace("{mes_corto}",codigo).replace("{año}",anio)

    # Obtenemos Excel
    df = obtener_df_url(url)

    # Si no encuentra dataframe
    if df is None:
        return None

    df = limpiar_filas_columnas_vacias_df(df)
    df.reset_index(inplace=True, drop=True)

    # Resetear los nombres de las columnas a 'col_1', 'col_2', 'col_3'
    df.columns = [f'col_{i+1}' for i in range(df.shape[1])]

    # Lista de subcadenas a buscar para eliminar las filas
    subcadenas_a_eliminar = ['Nota: ', '00:00:00', '(En miles de soles)', 'http://intranet1',
                            'Actualizado al', '(En miles de nuevos soles)', 'Actualizado el',
                            'Las definiciones de los tipos','A partir de','En el marco de la Emergencia',
                            'de la declaratoria de','presentaban más de','Actualizado',
                            'Indicadores Financieros por','Los valores anualizados se obtienen',
                            'Un crédito se considera vencido cuando','y de consumo, se considera',
                            'declaratoria de emergencia', 'más de 15 días calendario',
                            'La información del patrimonio efectivo']

    df = eliminar_filas_df_by_list(df, 'col_1', subcadenas_a_eliminar)
    df.reset_index(inplace=True, drop=True)

    # Eliminamos columnas repetidas donde se encuentre la palabra Activo
    columna_inicio_index = 1
    valor_a_eliminar = "SOLVENCIA"

    df = eliminar_columnas_df_by_value(df, valor_a_eliminar, columna_inicio_index)
    df.reset_index(inplace=True, drop=True)

    # Reemplazar el carácter '*' con un valor vacío en la columna 'col1'
    df['col_1'] = df['col_1'].str.replace('*', '', regex=False).apply(str.strip)

    # Para crear la columna Grupo se usa np.where() para asignar los valores según las condiciones
    df['Grupo'] = np.where(df['col_1']=='SOLVENCIA', 'SOLVENCIA',
                    np.where(df['col_1']=='CALIDAD DE ACTIVOS', 'CALIDAD DE ACTIVOS', 
                                np.where(df['col_1']=='EFICIENCIA Y GESTIÓN', 'EFICIENCIA Y GESTIÓN', 
                                        np.where(df['col_1']=='RENTABILIDAD', 'RENTABILIDAD', 
                                                np.where(df['col_1']=='LIQUIDEZ', 'LIQUIDEZ', 
                                                        np.where(df['col_1']=='POSICIÓN EN MONEDA EXTRANJERA', 'POSICIÓN EN MONEDA EXTRANJERA', None))))))

    # Rellenar los valores nulos en la columna 'Grupo' hacia adelante
    df['Grupo'] = df['Grupo'].ffill()

    # Eliminamos filas que aún quedan con la Lista de subcadenas a buscar para eliminar las filas
    lista_a_eliminar = ['SOLVENCIA', 'CALIDAD DE ACTIVOS', 'EFICIENCIA Y GESTIÓN', 'RENTABILIDAD', 
                            'LIQUIDEZ', 'POSICIÓN EN MONEDA EXTRANJERA']

    # Eliminar las filas donde la columna es igual a cualquiera de los valores en la lista
    df = df[~df['col_1'].isin(lista_a_eliminar)]

    df.reset_index(inplace=True, drop=True)

    # Colocamos el nombre de la 1er fila para la 1er columna y ultima
    df.iloc[0,0] = 'Indicador'
    df.iloc[0,-1] = 'Grupo'

    # Reemplazar los saltos de línea '\n' por un espacio en blanco ' ' en todas las columnas de la fila 1
    df.iloc[0] = df.iloc[0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) else x)

    # Reemplazar los doble espacio '  ' por un espacio en blanco ' ' en todas las columnas de la fila 1
    df.iloc[0] = df.iloc[0].apply(lambda x: x.replace('  ', ' ') if isinstance(x, str) else x)

    # Colocamos el nombre de la 1er fila para la 1er columna y ultima
    df = establecer_fila_como_nombre_columnas_df(df)

    df.reset_index(inplace=True)

    # Usar pd.melt() para convertir banco1 y banco2 en formato largo
    df = pd.melt(df, id_vars=['index', 'Indicador', 'Grupo'], value_vars=df.columns.difference(['index', 'Indicador', 'Grupo']),
                    var_name='Entidad', value_name='Valor')

    df['Valor'] = pd.to_numeric(df['Valor'])

    df['AnioMes'] = ult_dia_mes

    # Mapeo entre los nombres de las columnas del DataFrame y los de la tabla en SQL
    mapeo_columnas = {
        'AnioMes': 'dFecha',
        'index': 'nOrdenItem',
        'Grupo': 'cGrupo',
        'Indicador': 'cIndicador',
        'Entidad': 'cEntidad',
        'Valor': 'nValor'
    }

    df = df.rename(columns=mapeo_columnas)
    
    return df

# Créditos Directos según Situación: notebook5
def generarReporteSituacion(anio, mes, reporte=5):
    # Último deía del mes
    ult_dia_mes = obtener_ultimo_dia_mes(anio, mes)
    
    # Codigo de 2 dígitos
    codigo = obtener_codigo_mes(mes)
    
    # Obtenemos URL del reporte
    url = obtener_url_reporte(reporte)
    url = url.replace("{mes_largo}",mes).replace("{mes_corto}",codigo).replace("{año}",anio)
    
    # Obtenemos Excel
    df = obtener_df_url(url)

    # Si no encuentra dataframe
    if df is None:
        return None

    df = limpiar_filas_columnas_vacias_df(df)
    df.reset_index(inplace=True, drop=True)
    
    # Resetear los nombres de las columnas a 'col_1', 'col_2', 'col_3'
    df.columns = [f'col_{i+1}' for i in range(df.shape[1])]
    
    # Lista de subcadenas a buscar para eliminar las filas
    subcadenas_a_eliminar = ['Nota: ', '00:00:00', '(En miles de soles)', 'http://intranet1',
                            'Actualizado al', '(En miles de nuevos soles)', 'Actualizado el',
                            'Las definiciones de los tipos','A partir de','En el marco de la Emergencia',
                            'de la declaratoria de','presentaban más de','Actualizado',
                            'Indicadores Financieros por','Los valores anualizados se obtienen',
                            'Un crédito se considera vencido cuando','y de consumo, se considera',
                            'declaratoria de emergencia', 'más de 15 días calendario',
                            'La información del patrimonio efectivo']
    
    df = eliminar_filas_df_by_list(df, 'col_1', subcadenas_a_eliminar)
    df.reset_index(inplace=True, drop=True)
    
    # Reemplazar los saltos de línea '\n' por un espacio en blanco ' ' en todas las columnas de la fila 1
    df.iloc[0] = df.iloc[0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) else x)
    
    # Reemplazar los doble espacio '  ' por un espacio en blanco ' ' en todas las columnas de la fila 1
    df.iloc[0] = df.iloc[0].apply(lambda x: x.replace('  ', ' ') if isinstance(x, str) else x)
    
    # Seleccionar la fila que deseas corregir, en este caso, por ejemplo, la fila 0
    fila_a_corregir = 0
    
    # Aplicar el método 'ffill' para rellenar los valores nulos hacia adelante
    df.loc[fila_a_corregir] = df.loc[fila_a_corregir].ffill()
    
    # Concatenar las filas 0 y 1
    df.iloc[0,1:3] = df.iloc[0,1:3].fillna('').astype(str) + " " + df.iloc[1,1:3].fillna('').astype(str)
    
    # Eliminar la segunda fila después de la concatenación
    df = df.drop(index=1)
    
    df.reset_index(inplace=True, drop=True)
    
    # Colocamos el nombre de la 1er fila para la 1er columna y ultima
    df = establecer_fila_como_nombre_columnas_df(df)
    
    df.reset_index(inplace=True)

    # Usar pd.melt() para convertir banco1 y banco2 en formato largo
    df = pd.melt(df, id_vars=['index', 'Empresas'], value_vars=df.columns.difference(['index', 'Empresas']),
                var_name='Situación', value_name='Valor')

    df['Valor'] = pd.to_numeric(df['Valor'])

    df['AnioMes'] = ult_dia_mes

    # Mapeo entre los nombres de las columnas del DataFrame y los de la tabla en SQL
    mapeo_columnas = {
        'AnioMes': 'dFecha',
        'index': 'nOrdenItem',
        'Situación': 'cSituacion',
        'Empresas': 'cEntidad',
        'Valor': 'nValor'
    }

    df = df.rename(columns=mapeo_columnas)
    
    return df

# Estructura del Activo: notebook6
def generarReporteEstructuraActivo(anio, mes, reporte=6):
    # Último deía del mes
    ult_dia_mes = obtener_ultimo_dia_mes(anio, mes)

    # Codigo de 2 dígitos
    codigo = obtener_codigo_mes(mes)

    # Obtenemos URL del reporte
    url = obtener_url_reporte(reporte)
    url = url.replace("{mes_largo}",mes).replace("{mes_corto}",codigo).replace("{año}",anio)

    # Obtenemos Excel
    df = obtener_df_url(url)

    # Si no encuentra dataframe
    if df is None:
        return None

    df = limpiar_filas_columnas_vacias_df(df)
    df.reset_index(inplace=True, drop=True)

    # Resetear los nombres de las columnas a 'col_1', 'col_2', 'col_3'
    df.columns = [f'col_{i+1}' for i in range(df.shape[1])]

    # Lista de subcadenas a buscar para eliminar las filas
    subcadenas_a_eliminar = ['Nota: ', '00:00:00', '(En miles de soles)', 'http://intranet1',
                            'Actualizado al', '(En miles de nuevos soles)', 'Actualizado el',
                            'Las definiciones de los tipos','A partir de','En el marco de la Emergencia',
                            'de la declaratoria de','presentaban más de','Actualizado',
                            'Indicadores Financieros por','Los valores anualizados se obtienen',
                            'Un crédito se considera vencido cuando','y de consumo, se considera',
                            'declaratoria de emergencia', 'más de 15 días calendario',
                            'La información del patrimonio efectivo','(En porcentaje)',
                            'NOTA: Información obtenida del','Incluye activos no corrientes',
                            'Incluye fondos interbancarios']

    df = eliminar_filas_df_by_list(df, 'col_1', subcadenas_a_eliminar)
    df.reset_index(inplace=True, drop=True)

    # Reemplazar los saltos de línea '\n' por un espacio en blanco ' ' en todas las columnas de la fila 1
    df.iloc[0] = df.iloc[0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) else x)

    # Reemplazar los doble espacio '  ' por un espacio en blanco ' ' en todas las columnas de la fila 1
    df.iloc[0] = df.iloc[0].apply(lambda x: x.replace('  ', ' ') if isinstance(x, str) else x)

    # Reemplazar etiquetas 1/ y 2/ por un vacio '' en todas las columnas de la fila 1
    df.iloc[0] = df.iloc[0].apply(lambda x: x.replace('1/', '') if isinstance(x, str) else x)
    df.iloc[0] = df.iloc[0].apply(lambda x: x.replace('2/', '') if isinstance(x, str) else x)

    # Colocamos el nombre de la 1er fila para la 1er columna y ultima
    df = establecer_fila_como_nombre_columnas_df(df)

    df.reset_index(inplace=True)

    # Usar pd.melt() para convertir banco1 y banco2 en formato largo
    df = pd.melt(df, id_vars=['index', 'Empresas'], value_vars=df.columns.difference(['index', 'Empresas']),
                    var_name='Activo', value_name='Valor')

    df['Valor'] = pd.to_numeric(df['Valor'])

    df['AnioMes'] = ult_dia_mes

    # Mapeo entre los nombres de las columnas del DataFrame y los de la tabla en SQL
    mapeo_columnas = {
        'AnioMes': 'dFecha',
        'index': 'nOrdenItem',
        'Activo': 'cActivo',
        'Empresas': 'cEntidad',
        'Valor': 'nValor'
    }

    df = df.rename(columns=mapeo_columnas)
    
    return df

# Créditos Directos y Depósitos por Oficinas: notebook7
def generarReporteCreditosDepositosPorOficinas(anio, mes, reporte=7):
    # Último deía del mes
    ult_dia_mes = obtener_ultimo_dia_mes(anio, mes)

    # Codigo de 2 dígitos
    codigo = obtener_codigo_mes(mes)

    # Obtenemos URL del reporte
    url = obtener_url_reporte(reporte)
    url = url.replace("{mes_largo}",mes).replace("{mes_corto}",codigo).replace("{año}",anio)

    # Obtenemos Excel
    df = obtener_df_url(url)

    # Si no encuentra dataframe
    if df is None:
        return None

    df = limpiar_filas_columnas_vacias_df(df)
    df.reset_index(inplace=True, drop=True)

    # Resetear los nombres de las columnas a 'col_1', 'col_2', 'col_3'
    df.columns = [f'col_{i+1}' for i in range(df.shape[1])]

    # Lista de subcadenas a buscar para eliminar las filas
    subcadenas_a_eliminar = ['Nota: ', '00:00:00', '(En miles de soles)', 'http://intranet1',
                            'Actualizado al', '(En miles de nuevos soles)', 'Actualizado el',
                            'Las definiciones de los tipos','A partir de','En el marco de la Emergencia',
                            'de la declaratoria de','presentaban más de','Actualizado',
                            'Indicadores Financieros por','Los valores anualizados se obtienen',
                            'Un crédito se considera vencido cuando','y de consumo, se considera',
                            'declaratoria de emergencia', 'más de 15 días calendario',
                            'La información del patrimonio efectivo','(En porcentaje)',
                            'NOTA: Información obtenida del','Incluye activos no corrientes',
                            'Incluye fondos interbancarios']

    df = eliminar_filas_df_by_list(df, 'col_1', subcadenas_a_eliminar)
    df.reset_index(inplace=True, drop=True)

    # Seleccionar la fila que deseas corregir, en este caso, por ejemplo, la fila 0
    fila_a_corregir = 0

    # Aplicar el método 'ffill' para rellenar los valores nulos hacia adelante
    df.loc[fila_a_corregir] = df.loc[fila_a_corregir].ffill()

    # Reemplazar el valor 'nan' en la columna 'col2' con NaN
    df['col_1'] = df['col_1'].replace('nan', np.nan)

    # Concatenar las filas 0 y 1
    df.loc[0] = df.loc[0].fillna('').astype(str) + " " + df.loc[1].fillna('').astype(str)

    # Eliminar la segunda fila después de la concatenación
    df = df.drop(index=1)

    # Rellenar los valores nulos en la columna 'Grupo' hacia adelante
    df['col_1'] = df['col_1'].ffill()
    df['col_2'] = df['col_2'].ffill()
    df['col_3'] = df['col_3'].ffill()
    df['col_4'] = df['col_4'].ffill()

    # Colocamos el nombre de la 1er fila para la 1er columna y ultima
    df = establecer_fila_como_nombre_columnas_df(df)

    df.reset_index(inplace=True)

    # Cambiamos el tipo de dato
    df['Depósitos de Ahorro MN'] = pd.to_numeric(df['Depósitos de Ahorro MN'])
    df['Depósitos de Ahorro ME'] = pd.to_numeric(df['Depósitos de Ahorro ME'])
    df['Depósitos de Ahorro Total'] = pd.to_numeric(df['Depósitos de Ahorro Total'])
    df['Depósitos a Plazo MN'] = pd.to_numeric(df['Depósitos a Plazo MN'])
    df['Depósitos a Plazo ME'] = pd.to_numeric(df['Depósitos a Plazo ME'])
    df['Depósitos a Plazo Total'] = pd.to_numeric(df['Depósitos a Plazo Total'])
    df['Total Depósitos'] = pd.to_numeric(df['Total Depósitos'])
    df['Créditos Directos MN'] = pd.to_numeric(df['Créditos Directos MN'])
    df['Créditos Directos ME'] = pd.to_numeric(df['Créditos Directos ME'])
    df['Créditos Directos Total'] = pd.to_numeric(df['Créditos Directos Total'])

    df['AnioMes'] = ult_dia_mes

    # Mapeo entre los nombres de las columnas del DataFrame y los de la tabla en SQL
    mapeo_columnas = {
        'AnioMes': 'dFecha',
        'index': 'nOrdenItem',
        'Empresa': 'cEntidad',
        'Ubicación Departamento': 'cDepartamento',
        'Ubicación Provincia': 'cProvincia',
        'Ubicación Distrito': 'cDistrito',
        'Código de oficina': 'cCodOficina',
        'Depósitos de Ahorro MN': 'nDepositosAhorroMN',
        'Depósitos de Ahorro ME': 'nDepositosAhorroME',
        'Depósitos de Ahorro Total': 'nDepositosAhorroTotal',
        'Depósitos a Plazo MN': 'nDepositosPlazoMN',
        'Depósitos a Plazo ME': 'nDepositosPlazoME',
        'Depósitos a Plazo Total': 'nDepositosPlazoTotal',
        'Total Depósitos': 'nTotalDepósitos',
        'Créditos Directos MN': 'nCreditosDirectosMN',
        'Créditos Directos ME': 'nCreditosDirectosME',
        'Créditos Directos Total': 'nCreditosDirectosTotal'
    }

    df = df.rename(columns=mapeo_columnas)
    
    return df

# Número de Personal: notebook8
def generarReporteNumeroPersonal(anio, mes, reporte=8):
    # Último deía del mes
    ult_dia_mes = obtener_ultimo_dia_mes(anio, mes)

    # Codigo de 2 dígitos
    codigo = obtener_codigo_mes(mes)

    # Obtenemos URL del reporte
    url = obtener_url_reporte(reporte)
    url = url.replace("{mes_largo}",mes).replace("{mes_corto}",codigo).replace("{año}",anio)

    # Obtenemos Excel
    df = obtener_df_url(url)
    df = limpiar_filas_columnas_vacias_df(df)
    df.reset_index(inplace=True, drop=True)

    # Resetear los nombres de las columnas a 'col_1', 'col_2', 'col_3'
    df.columns = [f'col_{i+1}' for i in range(df.shape[1])]

    # Lista de subcadenas a buscar para eliminar las filas
    subcadenas_a_eliminar = ['Nota: ', '00:00:00', '(En miles de soles)', 'http://intranet1',
                            'Actualizado al', '(En miles de nuevos soles)', 'Actualizado el',
                            'Las definiciones de los tipos','A partir de','En el marco de la Emergencia',
                            'de la declaratoria de','presentaban más de','Actualizado',
                            'Indicadores Financieros por','Los valores anualizados se obtienen',
                            'Un crédito se considera vencido cuando','y de consumo, se considera',
                            'declaratoria de emergencia', 'más de 15 días calendario',
                            'La información del patrimonio efectivo','(En porcentaje)',
                            'NOTA: Información obtenida del','Incluye activos no corrientes',
                            'Incluye fondos interbancarios', '(En número de personas)']

    df = eliminar_filas_df_by_list(df, 'col_1', subcadenas_a_eliminar)
    df.reset_index(inplace=True, drop=True)

    # Colocamos el nombre de la 1er fila para la 1er columna y ultima
    df = establecer_fila_como_nombre_columnas_df(df)

    df.reset_index(inplace=True)

    # Usar pd.melt() para convertir banco1 y banco2 en formato largo
    df = pd.melt(df, id_vars=['index', 'Empresas'], value_vars=df.columns.difference(['index', 'Empresas']),
                    var_name='CategoriaPersonal', value_name='Valor')

    df['Valor'] = pd.to_numeric(df['Valor'])

    df['AnioMes'] = ult_dia_mes

    # Mapeo entre los nombres de las columnas del DataFrame y los de la tabla en SQL
    mapeo_columnas = {
        'AnioMes': 'dFecha',
        'index': 'nOrdenItem',
        'Empresas': 'cEntidad',
        'CategoriaPersonal': 'cCategoriaPersonal',
        'Valor': 'nValor'
    }

    df = df.rename(columns=mapeo_columnas)

    return df