import os
import pandas as pd
import numpy as np

ruta_carpeta = r'C:\Users\Carla\Documents\analisis-horarios\datos_originales'
nombre_archivo = 'horario_aulas_limpieza.xlsx'
ruta_completa = os.path.join(ruta_carpeta, nombre_archivo)

#Abrimos el archivo Excel y agregamos una columna 'id' que contiene el índice de cada fila.
df= pd.read_excel(ruta_completa)
df['id'] = df.index 

#Limpiamos las columnas de texto: eliminamos espacios al inicio y al final, reemplazamos múltiples espacios por uno solo, convertimos a mayúsculas y reemplazamos cadenas vacías por NaN.
columnas_texto = df.select_dtypes(include=['object']).columns
for col in columnas_texto:
    df[col] = (df[col]
               .str.strip()                           
               .str.replace(r'\s+', ' ', regex=True)  
               .str.upper()                           
               .replace('', np.nan)                  
              )

# Convertimos las columnas de fecha y hora a tipo datetime.
df['HORAINICIO'] = pd.to_datetime(df['HORAINICIO'])
df['HORAFIN'] = pd.to_datetime(df['HORAFIN'])

# Calcular la diferencia en minutos
df['DURACIONMINUTOS'] = (df['HORAFIN'] - df['HORAINICIO']).dt.total_seconds() / 60

# Convertimos las columnas de hora a tipo time para facilitar la visualización.
df['HORAINICIO'] = df['HORAINICIO'].dt.time
df['HORAFIN'] = df['HORAFIN'].dt.time

#Convertimos la columna 'ANIO' a tipo numérico.
df['ANIO'] = pd.to_numeric(df['ANIO'], errors='coerce')

#Filtramos el DataFrame para obtener solo las filas correspondientes al año 2026, término '1S' y campus 'CAMPUS GUSTAVO GALINDO'.
campus_objetivo = 'CAMPUS GUSTAVO GALINDO'  
df_filtrado = df[(df['ANIO'] == 2026) & 
                 (df['TERMINO'] == '1S') & 
                 (df['CAMPUS'] == campus_objetivo)]

