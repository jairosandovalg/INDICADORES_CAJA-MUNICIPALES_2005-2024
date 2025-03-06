import pandas as pd
import os

os.chdir(r"D:\ANALISTA DE DATOS\PROYECTOS\INDICADORES_CM")
# Ruta de las carpetas
ruta_carpetas = r"D:\ANALISTA DE DATOS\PROYECTOS\INDICADORES_CM\MATERIAL"
carpetas = os.listdir(ruta_carpetas)

# Mapeo de meses en texto a números
meses_dict = {'en': '01', 'fe': '02', 'ma': '03', 'ab': '04', 'my': '05', 
              'jn': '06', 'jl': '07', 'ag': '08', 'se': '09', 'oc': '10', 
              'no': '11', 'di': '12'}

df = {}
datos_consolidados = []

for carpeta in carpetas:
    ruta_carpeta = os.path.join(ruta_carpetas, carpeta)
    archivos = os.listdir(ruta_carpeta)
    
    for archivo in archivos:
        ruta_archivo = os.path.join(ruta_carpeta, archivo)
        
        
        # Cargar archivo Excel
        df_mensual = pd.read_excel(ruta_archivo)

        # Eliminar filas con menos de 10 valores no nulos
        df_mensual.dropna(thresh=10, inplace=True)
        # Eliminar columnas con menos de 15 valores no nulos
        df_mensual.dropna(thresh=15, axis=1, inplace=True)
        # Reiniciar el índice
        df_mensual.reset_index(drop=True, inplace=True)

        # Asignar la primera fila como nombres de las columnas
        df_mensual.columns = df_mensual.iloc[0].astype(str).fillna('').str.replace('\u00A0', ' ')
        df_mensual = df_mensual[1:].reset_index(drop=True)
        

        # Reemplazar el nombre de la primera columna si es NaN
        df_mensual.columns.values[0] = "Indicadores"
        # Establecer la primera columna como índice
        df_mensual.set_index("Indicadores", inplace=True)
        # Filtrar y conservar solo las columnas que tienen nombres válidos (no NaN)
        df_mensual = df_mensual.loc[:, df_mensual.columns.notna() & (df_mensual.columns.astype(str) != 'nan')]
        # Eliminar índices vacíos o NaN
        df_mensual = df_mensual[~df_mensual.index.isna() & (df_mensual.index.astype(str) != 'nan') & (df_mensual.index != '')]

        
        # Limpiar nombres de columnas
        df_mensual.columns = (df_mensual.columns.astype(str).fillna('')
                              .str.replace('\u00A0', ' ')  # Reemplazar espacios no separables
                              .str.replace(r'\s+', ' ', regex=True)  # Reducir múltiples espacios a uno
                              .str.replace(',', '')  # Eliminar comas
                              .str.replace(r'\*', '', regex=True)  # Eliminar asteriscos
                              .str.replace(r'\(.*?\)', '', regex=True) #Elimina cada paréntesis con su contenido
                              .str.replace(r'\s*1/', '', regex=True)  # Eliminar "1/" con su espacio antes
                              .str.strip())  # Eliminar espacios al inicio y final
        
        # Limpiar nombres de filas (índices)
        df_mensual.index = (df_mensual.index.astype(str).fillna('')
                                     .str.replace('\u00A0', ' ')  # Reemplazar espacios no separables
                                     .str.replace(r'\s+', ' ', regex=True)  # Reducir múltiples espacios a uno
                                     .str.replace(',', '')  # Eliminar comas
                                     .str.replace(r'\*', '', regex=True)  # Eliminar asteriscos
                                     .str.replace(r'\(.*?\)', '', regex=True) #Elimina cada paréntesis con su contenido
                                     .str.replace(r'\s+\'', "'", regex=True)  # Eliminar espacios antes de la última comilla
                                     .str.replace(r'\s*1/', '', regex=True)  # Eliminar "1/" y su espacio antes
                                     .str.replace(r'\s*promedio del mes\s*', '', regex=True)  # Eliminar "promedio del mes"
                                     .str.replace(r'\s*al\s*\d{1,2}/\d{1,2}/\d{4}', '', regex=True)  # Eliminar fechas dd/mm/aaaa
                                     .str.replace(r'\s*al\s*\d{3,4}/\d{4}', '', regex=True)  # Eliminar formatos tipo 307/2009
                                     .str.replace(r'\s*/\s*', ' ', regex=True)  # Eliminar "/" con espacios alrededor
                                     .str.replace(r'\s*\d+\s*$', '', regex=True)  # Eliminar números al final de la cadena
                                     .str.replace(r'\s+\.\s*$', '', regex=True)  # Eliminar punto y espacios finales
                                     # Unificar nombres de indicadores financieros
                                     .str.replace(
                                         r".*Ingresos Financieros.*",
                                         r"Ingresos Financieros Anualizados / Activo Productivo Promedio %",
                                         regex=True
                                     )
                                     .str.replace(
                                         r"Provisiones Cartera Atrasada %",
                                         "Provisiones Créditos Atrasados %",
                                         regex=True
                                     )

                                    .str.replace(
                                        r"Créditos Atrasados criterio SBS Créditos Directos|"
                                        r"Créditos Atrasados Créditos Directos %|"
                                        r"Cartera Atrasada Créditos Directos %",
                                        "Créditos Atrasados / Créditos Directos",
                                        regex=True
                                    )
                                    .str.replace(
                                        r"Créditos Atrasados MN \(criterio SBS\)\*\* / Créditos Directos MN",
                                        "Créditos Atrasados MN / Créditos Directos MN",
                                        regex=True
                                    )
                                     .str.replace(
                                        r"Créditos Atrasados M.E. Créditos Directos M.E. %|"
                                        r"Cartera Atrasada M.E. Créditos Directos M.E. %|"
                                        r"Créditos Atrasados ME criterio SBS Créditos Directos ME",
                                        "Créditos Atrasados ME / Créditos Directos ME",
                                        regex=True
                                    )
                                    .str.replace(
                                        r"Créditos Atrasados M.N. Créditos Directos M.N. %|"
                                        r"Créditos Atrasados MN criterio SBS Créditos Directos MN|"
                                        r"Cartera Atrasada M.N. Créditos Directos M.N. %",
                                        "Créditos Atrasados MN / Créditos Directos MN",
                                        regex=True
                                    )
                                    .str.replace(
                                        r"Cartera de Alto Riesgo Créditos Directos %",
                                        "Cartera de Alto Riesgo / Créditos Directos (%)",
                                        regex=True
                                    )
                                    .str.replace(
                                        r".*Gastos de Operación.*",
                                        r"Gastos de Operación Anualizados / Margen Financiero Total Anualizado (%)",
                                        regex=True
                                    )
                                    .str.replace(
                                        r'.*Gastos de Administración.*',
                                        r'Gastos de Administración Anualizados/ Créditos Directos e Indirectos Promedio (%)', 
                                        regex = True
                                
                                   ) 
                                    .str.strip())


        # Identificar el mes en el nombre del archivo
        mes = next((m for m in meses_dict if m in archivo.lower()), None)       

        periodo = pd.to_datetime(f"{carpeta}-{meses_dict[mes]}", format="%Y-%m").strftime("%Y-%m")
        df[periodo] = df_mensual

        # Extraer valores
        for entidad in df_mensual.columns:
            valores = df_mensual.reindex(df_mensual.index)[entidad]
            for indicador, valor in valores.items():  
                datos_consolidados.append([entidad, periodo, indicador, valor])
        
        # Crear DataFrame consolidado
df_final = pd.DataFrame(datos_consolidados, columns=["Entidad", "Periodo", "Indicador", "Monto"])
df_final["Monto"] = pd.to_numeric(df_final["Monto"], errors="coerce")

df_final.to_excel("INDICADORES_CM_2005-2024.xlsx", sheet_name="IND_CM")

