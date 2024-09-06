import streamlit as st
from reportes import *

st.title("Reportes SBS")

# Configura tu conexión
server = 'SRVPVMRIESG02\\SQLSERVERUAR'
database = 'UAR'

# Ejemplo de ejecución de una consulta
query = 'SELECT * FROM SF.URLs WHERE nTipo=3 AND nIdReporte = {numId}'

# Diccionario del proceso
reportes = {
    "Balance General": 1,
    "Ganacias y Perdidas": 1,
    "Sector Economico Empresariales": 2,
    "Tipo de Credito por Situacion": 3,
    "Indicadores Financieros": 4,
    "Situacion": 5,
    "Estructura del Activo": 6,
    "Creditos y Depositos por Oficinas": 7,
    "Número de personal": 8,
}

lista_meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Setiembre', 'Octubre', 'Noviembre', 'Diciembre']

# anio = st.text_input("Ingrese el año en formato yyyy", placeholder="Ejemplo: 2024")
anio = st.number_input("Pick a number", min_value=2015, max_value=2024, value=2024, step=1)

meses = st.multiselect("Elija el mes a procesar", options=lista_meses, max_selections=2)
procesos = st.multiselect("Elija el reporte a procesar", options=reportes.keys(), max_selections=2)

procesar_btn = st.button("Procesar", type="primary")

anio = str(anio)

if procesar_btn:
    for mes in meses:
        for proceso in procesos:
            if proceso == "Balance General":
                numId = reportes[proceso]
                query = f'SELECT * FROM SF.URLs WHERE nTipo=3 AND nIdReporte = {numId}'
                df = generarReporteBalanceGeneral(server,database,query,anio,mes)
                # grabar_df_sql(server, database, df, 'BalanceGeneral', 'SF')
                st.write(df)
            elif proceso == "Ganacias y Perdidas":
                numId = reportes[proceso]
                query = f'SELECT * FROM SF.URLs WHERE nTipo=3 AND nIdReporte = {numId}'
                df = generarReporteEstadoGananciasPerdidas(server,database,query,anio,mes)
                # grabar_df_sql(server, database, df, 'GananciasPerdidas', 'SF')

            elif proceso == "Sector Economico Empresariales":
                numId = reportes[proceso]
                query = f'SELECT * FROM SF.URLs WHERE nTipo=3 AND nIdReporte = {numId}'
                df = generarReporteSectorEconomicoEmpresariales(server,database,query,anio,mes)
                # grabar_df_sql(server, database, df, 'SectorEconomicoEmpresariales', 'SF')

            elif proceso == "Tipo de Credito por Situacion":
                numId = reportes[proceso]
                query = f'SELECT * FROM SF.URLs WHERE nTipo=3 AND nIdReporte = {numId}'
                df = generarReporteTipoCreditoSituacion(server,database,query,anio,mes)
                # grabar_df_sql(server, database, df, 'TipoCreditoSituacion', 'SF')
                
            elif proceso == "Indicadores Financieros":
                numId = reportes[proceso]
                query = f'SELECT * FROM SF.URLs WHERE nTipo=3 AND nIdReporte = {numId}'
                df = generarReporteIndicadoresFinancieros(server,database,query,anio,mes)
                # grabar_df_sql(server, database, df, 'IndicadoresFinancieros', 'SF')

            elif proceso == "Situacion":
                numId = reportes[proceso]
                query = f'SELECT * FROM SF.URLs WHERE nTipo=3 AND nIdReporte = {numId}'
                df = generarReporteSituacion(server,database,query,anio,mes)
                # grabar_df_sql(server, database, df, 'Situacion', 'SF')

            elif proceso == "Estructura del Activo":
                numId = reportes[proceso]
                query = f'SELECT * FROM SF.URLs WHERE nTipo=3 AND nIdReporte = {numId}'
                df = generarReporteEstructuraActivo(server,database,query,anio,mes)
                # grabar_df_sql(server, database, df, 'EstructuraActivo', 'SF')

            elif proceso == "Creditos y Depositos por Oficinas":
                numId = reportes[proceso]
                query = f'SELECT * FROM SF.URLs WHERE nTipo=3 AND nIdReporte = {numId}'
                df = generarReporteCreditosDepositosPorOficinas(server,database,query,anio,mes)
                # grabar_df_sql(server, database, df, 'CreditosDepositosPorOficina', 'SF')
            
            elif proceso == "Número de personal":
                numId = reportes[proceso]
                query = f'SELECT * FROM SF.URLs WHERE nTipo=3 AND nIdReporte = {numId}'
                df = generarReporteNumeroPersonal(server,database,query,anio,mes)
                # grabar_df_sql(server, database, df, 'CreditosDepositosPorOficina', 'SF')
                st.write(df)