import os
import unicodedata
from datetime import datetime, time

import numpy as np
import pandas as pd


# ==========================================================
# 1. RUTAS Y PARÁMETROS
# ==========================================================

ruta_proyecto = (
    r'/workspaces/'
    r'Analisis-de-horarios-de-clase-y-examenes-planificados'
)

ruta_entrada = os.path.join(
    ruta_proyecto,
    'datos_originales',
    'horario_examenes_limpieza.xlsx'
)

ruta_salida = os.path.join(
    ruta_proyecto,
    'datos_limpios',
    'resultado_limpieza_examenes.xlsx'
)

anio_objetivo = 2026
termino_objetivo = '1S'
campus_objetivo = 'CAMPUS GUSTAVO GALINDO'

# Solo se utiliza como control.
# Una duración superior a diez horas se marca como atípica.
duracion_maxima_minutos = 600


# ==========================================================
# 2. COLUMNAS OBLIGATORIAS
# ==========================================================

columnas_obligatorias = {
    'ANIO',
    'TERMINO',
    'FECHA',
    'EXAMEN',
    'DIA',
    'HORAINICIO',
    'HORAFIN',
    'CODIGOMATERIA',
    'PARALELO',
    'NUMREGISTRADOS',
    'ESVIRTUAL',
    'AULA',
    'BLOQUE',
    'CAMPUS',
    'REFERENCIA'
}


# ==========================================================
# 3. FUNCIONES AUXILIARES
# ==========================================================

def quitar_tildes(valor):
    """
    Convierte:
    MIÉRCOLES -> MIERCOLES
    SÁBADO -> SABADO
    """

    if pd.isna(valor):
        return pd.NA

    texto = unicodedata.normalize(
        'NFKD',
        str(valor)
    )

    return ''.join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )


def hora_a_minutos(valor):
    """
    Convierte una hora de Excel, datetime, time o texto
    en minutos transcurridos desde las 00:00.

    Ejemplo:
    09:00 -> 540
    11:30 -> 690
    """

    if pd.isna(valor):
        return pd.NA

    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.hour * 60 + valor.minute

    if isinstance(valor, time):
        return valor.hour * 60 + valor.minute

    if isinstance(
        valor,
        (
            int,
            float,
            np.integer,
            np.floating
        )
    ):
        numero = float(valor)

        # Fracción de día de Excel.
        if 0 <= numero < 1:
            return round(numero * 1440) % 1440

        # Fecha serial de Excel con una fracción horaria.
        return round((numero % 1) * 1440) % 1440

    convertido = pd.to_datetime(
        str(valor).strip(),
        errors='coerce'
    )

    if pd.isna(convertido):
        return pd.NA

    return (
        convertido.hour * 60
        + convertido.minute
    )


def minutos_a_hora(valor):
    """
    Convierte minutos desde medianoche a un objeto time.

    Ejemplo:
    540 -> 09:00
    """

    if pd.isna(valor):
        return pd.NA

    horas, minutos = divmod(
        int(valor) % 1440,
        60
    )

    return time(horas, minutos)


# ==========================================================
# 4. COMPROBAR Y CARGAR EL ARCHIVO
# ==========================================================

if not os.path.exists(ruta_entrada):
    raise FileNotFoundError(
        f'No se encontró el archivo:\n{ruta_entrada}'
    )

df = pd.read_excel(ruta_entrada)

if df.empty:
    raise ValueError(
        'El archivo de exámenes está vacío.'
    )

columnas_faltantes = (
    columnas_obligatorias
    - set(df.columns)
)

if columnas_faltantes:
    raise ValueError(
        'Faltan las siguientes columnas obligatorias: '
        + ', '.join(
            sorted(columnas_faltantes)
        )
    )

# La fila 1 corresponde a los encabezados.
df.insert(
    0,
    'ID_FILA_ORIGINAL',
    df.index + 2
)

cantidad_original = len(df)


# ==========================================================
# 5. LIMPIAR COLUMNAS DE TEXTO
# ==========================================================

columnas_texto = df.select_dtypes(
    include=['object', 'string']
).columns

for columna in columnas_texto:
    df[columna] = (
        df[columna]
        .astype('string')
        .str.strip()
        .str.replace(
            r'\s+',
            ' ',
            regex=True
        )
        .str.upper()
        .replace('', pd.NA)
    )


# ==========================================================
# 6. CONSERVAR VALORES REPORTADOS
# ==========================================================

df['FECHA_REPORTADA'] = df['FECHA']

df['HORA_INICIO_REPORTADA'] = (
    df['HORAINICIO']
)

df['HORA_FIN_REPORTADA'] = (
    df['HORAFIN']
)

df['NUMREGISTRADOS_REPORTADO'] = (
    df['NUMREGISTRADOS']
)


# ==========================================================
# 7. CONVERTIR COLUMNAS NUMÉRICAS
# ==========================================================

df['ANIO'] = pd.to_numeric(
    df['ANIO'],
    errors='coerce'
).astype('Int64')

df['PARALELO'] = pd.to_numeric(
    df['PARALELO'],
    errors='coerce'
).astype('Int64')

df['NUMREGISTRADOS'] = pd.to_numeric(
    df['NUMREGISTRADOS'],
    errors='coerce'
)


# ==========================================================
# 8. CONVERTIR FECHA
# ==========================================================

df['FECHA'] = pd.to_datetime(
    df['FECHA'],
    errors='coerce',
    dayfirst=True
)


# ==========================================================
# 9. CONVERTIR HORAS
# ==========================================================

df['MINUTO_INICIO'] = (
    df['HORAINICIO']
    .map(hora_a_minutos)
    .astype('Int64')
)

df['MINUTO_FIN'] = (
    df['HORAFIN']
    .map(hora_a_minutos)
    .astype('Int64')
)

df['HORAINICIO'] = (
    df['MINUTO_INICIO']
    .map(minutos_a_hora)
)

df['HORAFIN'] = (
    df['MINUTO_FIN']
    .map(minutos_a_hora)
)

df['DURACIONMINUTOS'] = (
    df['MINUTO_FIN']
    - df['MINUTO_INICIO']
).astype('Int64')


# ==========================================================
# 10. FILTRAR PERIODO Y CAMPUS
# ==========================================================

df_filtrado = df[
    (df['ANIO'] == anio_objetivo)
    & (df['TERMINO'] == termino_objetivo)
    & (df['CAMPUS'] == campus_objetivo)
].copy().reset_index(drop=True)

print(
    f'Registros originales: {cantidad_original}'
)

print(
    'Registros del periodo y campus objetivo: '
    f'{len(df_filtrado)}'
)


# ==========================================================
# 11. VALIDAR FECHA Y DÍA
# ==========================================================

df_filtrado['FLAG_FECHA_INVALIDA'] = (
    df_filtrado['FECHA']
    .isna()
    .astype(int)
)

df_filtrado['DIA_REPORTADO'] = (
    df_filtrado['DIA']
    .map(quitar_tildes)
    .astype('string')
    .str.upper()
)

dias_semana = {
    0: 'LUNES',
    1: 'MARTES',
    2: 'MIERCOLES',
    3: 'JUEVES',
    4: 'VIERNES',
    5: 'SABADO',
    6: 'DOMINGO'
}

df_filtrado['DIA_CALCULADO'] = (
    df_filtrado['FECHA']
    .dt.dayofweek
    .map(dias_semana)
    .astype('string')
)

df_filtrado['FLAG_DIA_FALTANTE'] = (
    df_filtrado['DIA_REPORTADO']
    .isna()
    .astype(int)
)

condicion_dia_inconsistente = (
    df_filtrado['DIA_REPORTADO'].notna()
    & df_filtrado['DIA_CALCULADO'].notna()
    & (
        df_filtrado['DIA_REPORTADO']
        != df_filtrado['DIA_CALCULADO']
    )
)

df_filtrado['FLAG_DIA_INCONSISTENTE'] = (
    condicion_dia_inconsistente
    .fillna(False)
    .astype(int)
)


# ==========================================================
# 12. VALIDAR TIPO DE EXAMEN
# ==========================================================

tipos_examen_validos = [
    'PARCIAL',
    'FINAL',
    'MEJORAMIENTO'
]

condicion_tipo_examen_invalido = (
    df_filtrado['EXAMEN'].isna()
    |
    ~df_filtrado['EXAMEN'].isin(
        tipos_examen_validos
    )
)

df_filtrado[
    'FLAG_TIPO_EXAMEN_INVALIDO'
] = condicion_tipo_examen_invalido.astype(int)


# ==========================================================
# 13. VALIDAR HORARIOS
# ==========================================================

# Solo se considera horario inválido cuando:
# - Falta la hora de inicio.
# - Falta la hora de fin.
# - La hora final es anterior a la inicial.
#
# Las horas iguales se conservan como una bandera
# independiente para revisión.

condicion_hora_invalida = (
    df_filtrado['MINUTO_INICIO'].isna()
    | df_filtrado['MINUTO_FIN'].isna()
    |
    (
        df_filtrado['MINUTO_FIN']
        < df_filtrado['MINUTO_INICIO']
    )
)

df_filtrado['FLAG_HORA_INVALIDA'] = (
    condicion_hora_invalida
    .fillna(False)
    .astype(int)
)

condicion_horas_iguales = (
    df_filtrado['MINUTO_INICIO'].notna()
    & df_filtrado['MINUTO_FIN'].notna()
    &
    (
        df_filtrado['MINUTO_INICIO']
        == df_filtrado['MINUTO_FIN']
    )
)

df_filtrado['FLAG_HORAS_IGUALES'] = (
    condicion_horas_iguales.astype(int)
)

condicion_duracion_atipica = (
    (df_filtrado['FLAG_HORA_INVALIDA'] == 0)
    &
    (
        df_filtrado['DURACIONMINUTOS']
        > duracion_maxima_minutos
    )
)

df_filtrado['FLAG_DURACION_ATIPICA'] = (
    condicion_duracion_atipica
    .fillna(False)
    .astype(int)
)


# ==========================================================
# 14. VALIDAR NÚMERO DE ESTUDIANTES
# ==========================================================

df_filtrado[
    'FLAG_REGISTRADOS_FALTANTE'
] = (
    df_filtrado['NUMREGISTRADOS']
    .isna()
    .astype(int)
)

df_filtrado[
    'FLAG_REGISTRADOS_NEGATIVO'
] = (
    df_filtrado['NUMREGISTRADOS']
    .lt(0)
    .fillna(False)
    .astype(int)
)

df_filtrado[
    'FLAG_REGISTRADOS_CERO'
] = (
    df_filtrado['NUMREGISTRADOS']
    .eq(0)
    .fillna(False)
    .astype(int)
)

# En exámenes no se heredan estudiantes
# de un paralelo teórico.
df_filtrado[
    'NUMREGISTRADOS_AJUSTADO'
] = (
    pd.to_numeric(
        df_filtrado['NUMREGISTRADOS'],
        errors='coerce'
    )
    .round()
    .astype('Int64')
)

df_filtrado[
    'ORIGEN_NUMREGISTRADOS'
] = 'ORIGINAL'


# ==========================================================
# 15. VALIDAR VIRTUALIDAD
# ==========================================================

condicion_esvirtual_no_resuelto = (
    df_filtrado['ESVIRTUAL'].isna()
    |
    ~df_filtrado['ESVIRTUAL'].isin(
        ['S', 'N']
    )
)

df_filtrado[
    'FLAG_ESVIRTUAL_NO_RESUELTO'
] = condicion_esvirtual_no_resuelto.astype(int)

# El registro aparece como virtual, pero:
# - Está en Gustavo Galindo.
# - Tiene aula.
# - Tiene bloque.
#
# Se envía a revisión y no se elimina.
condicion_virtual_en_campus_fisico = (
    (df_filtrado['ESVIRTUAL'] == 'S')
    & df_filtrado['AULA'].notna()
    & df_filtrado['BLOQUE'].notna()
)

df_filtrado[
    'FLAG_VIRTUAL_EN_CAMPUS_FISICO'
] = (
    condicion_virtual_en_campus_fisico
    .fillna(False)
    .astype(int)
)


# ==========================================================
# 16. VALIDAR UBICACIÓN
# ==========================================================

condicion_ubicacion_invalida = (
    df_filtrado['AULA'].isna()
    | df_filtrado['BLOQUE'].isna()
    |
    df_filtrado['AULA']
    .fillna('')
    .str.contains(
        'VIRTUAL',
        regex=False
    )
    |
    df_filtrado['BLOQUE']
    .fillna('')
    .str.contains(
        'VIRTUAL',
        regex=False
    )
)

df_filtrado[
    'FLAG_UBICACION_INVALIDA'
] = (
    condicion_ubicacion_invalida
    .fillna(False)
    .astype(int)
)

df_filtrado['ID_ESPACIO'] = (
    df_filtrado['BLOQUE']
    .fillna('SIN_BLOQUE')
    .astype(str)
    + '|'
    + df_filtrado['AULA']
    .fillna('SIN_AULA')
    .astype(str)
)


# ==========================================================
# 17. IDENTIFICAR EVENTOS FÍSICOS
# ==========================================================

# Es una clasificación provisional:
# ESVIRTUAL=N y tiene una ubicación válida.

condicion_evento_fisico = (
    (df_filtrado['ESVIRTUAL'] == 'N')
    & (
        df_filtrado[
            'FLAG_UBICACION_INVALIDA'
        ] == 0
    )
    & (
        df_filtrado[
            'FLAG_ESVIRTUAL_NO_RESUELTO'
        ] == 0
    )
)

df_filtrado['ES_EVENTO_FISICO'] = (
    condicion_evento_fisico
    .fillna(False)
    .astype(int)
)


# ==========================================================
# 18. DETECTAR EVENTOS REPETIDOS
# ==========================================================

columnas_evento = [
    'FECHA',
    'EXAMEN',
    'HORAINICIO',
    'HORAFIN',
    'CODIGOMATERIA',
    'PARALELO',
    'BLOQUE',
    'AULA'
]

df_filtrado['FLAG_EVENTO_REPETIDO'] = (
    df_filtrado
    .duplicated(
        subset=columnas_evento,
        keep=False
    )
    .astype(int)
)


# ==========================================================
# 19. DETECTAR DUPLICADOS EXACTOS
# ==========================================================

columnas_duplicado_exacto = [
    'ANIO',
    'TERMINO',
    'FECHA',
    'EXAMEN',
    'DIA',
    'HORAINICIO',
    'HORAFIN',
    'CODIGOMATERIA',
    'PARALELO',
    'NUMREGISTRADOS',
    'ESVIRTUAL',
    'ESELEARNING',
    'TIPOELEARNING',
    'ESHIBRIDO',
    'AULA',
    'BLOQUE',
    'CAMPUS',
    'REFERENCIA'
]

# Solo utilizamos las columnas que realmente existen.
columnas_duplicado_exacto = [
    columna
    for columna in columnas_duplicado_exacto
    if columna in df_filtrado.columns
]

df_filtrado['FLAG_DUPLICADO_EXACTO'] = (
    df_filtrado
    .duplicated(
        subset=columnas_duplicado_exacto,
        keep=False
    )
    .astype(int)
)


# ==========================================================
# 20. CLASIFICAR REGISTROS
# ==========================================================

condicion_invalido = (
    (df_filtrado['FLAG_FECHA_INVALIDA'] == 1)
    |
    (df_filtrado['FLAG_HORA_INVALIDA'] == 1)
    |
    (
        df_filtrado[
            'FLAG_REGISTRADOS_FALTANTE'
        ] == 1
    )
    |
    (
        df_filtrado[
            'FLAG_REGISTRADOS_NEGATIVO'
        ] == 1
    )
    |
    (
        df_filtrado[
            'FLAG_UBICACION_INVALIDA'
        ] == 1
    )
    |
    (
        df_filtrado[
            'FLAG_TIPO_EXAMEN_INVALIDO'
        ] == 1
    )
)

condicion_revisar = (
    ~condicion_invalido
    &
    (
        (
            df_filtrado[
                'FLAG_DIA_FALTANTE'
            ] == 1
        )
        |
        (
            df_filtrado[
                'FLAG_DIA_INCONSISTENTE'
            ] == 1
        )
        |
        (
            df_filtrado[
                'FLAG_HORAS_IGUALES'
            ] == 1
        )
        |
        (
            df_filtrado[
                'FLAG_DURACION_ATIPICA'
            ] == 1
        )
        |
        (
            df_filtrado[
                'FLAG_VIRTUAL_EN_CAMPUS_FISICO'
            ] == 1
        )
        |
        (
            df_filtrado[
                'FLAG_ESVIRTUAL_NO_RESUELTO'
            ] == 1
        )
        |
        (
            df_filtrado[
                'FLAG_REGISTRADOS_CERO'
            ] == 1
        )
        |
        (
            df_filtrado[
                'FLAG_EVENTO_REPETIDO'
            ] == 1
        )
        |
        (
            df_filtrado[
                'FLAG_DUPLICADO_EXACTO'
            ] == 1
        )
    )
)

df_filtrado['ESTADO_REGISTRO'] = 'VALIDO'

df_filtrado.loc[
    condicion_revisar,
    'ESTADO_REGISTRO'
] = 'REVISAR'

df_filtrado.loc[
    condicion_invalido,
    'ESTADO_REGISTRO'
] = 'INVALIDO'


# ==========================================================
# 21. CREAR MOTIVO DEL ESTADO
# ==========================================================

def obtener_motivo_estado(fila):

    motivos = []

    if fila['FLAG_FECHA_INVALIDA'] == 1:
        motivos.append(
            'FECHA INVALIDA'
        )

    if fila['FLAG_DIA_FALTANTE'] == 1:
        motivos.append(
            'DIA REPORTADO FALTANTE'
        )

    if fila['FLAG_DIA_INCONSISTENTE'] == 1:
        motivos.append(
            'DIA NO COINCIDE CON FECHA'
        )

    if fila['FLAG_TIPO_EXAMEN_INVALIDO'] == 1:
        motivos.append(
            'TIPO DE EXAMEN INVALIDO'
        )

    if fila['FLAG_HORA_INVALIDA'] == 1:
        motivos.append(
            'HORARIO INVALIDO'
        )

    if fila['FLAG_HORAS_IGUALES'] == 1:
        motivos.append(
            'HORAS DE INICIO Y FIN IGUALES'
        )

    if fila['FLAG_DURACION_ATIPICA'] == 1:
        motivos.append(
            'DURACION MAYOR A 5 HORAS'
        )

    if fila['FLAG_REGISTRADOS_FALTANTE'] == 1:
        motivos.append(
            'NUMERO DE REGISTRADOS FALTANTE'
        )

    if fila['FLAG_REGISTRADOS_NEGATIVO'] == 1:
        motivos.append(
            'NUMERO DE REGISTRADOS NEGATIVO'
        )

    if fila['FLAG_REGISTRADOS_CERO'] == 1:
        motivos.append(
            'NUMERO DE REGISTRADOS EN CERO'
        )

    if fila['FLAG_ESVIRTUAL_NO_RESUELTO'] == 1:
        motivos.append(
            'ESVIRTUAL NO RESUELTO'
        )

    if fila[
        'FLAG_VIRTUAL_EN_CAMPUS_FISICO'
    ] == 1:
        motivos.append(
            'ESVIRTUAL=S CON AULA Y BLOQUE FISICOS'
        )

    if fila['FLAG_UBICACION_INVALIDA'] == 1:
        motivos.append(
            'AULA O BLOQUE INVALIDO'
        )

    if fila['FLAG_EVENTO_REPETIDO'] == 1:
        motivos.append(
            'POSIBLE EVENTO REPETIDO'
        )

    if fila['FLAG_DUPLICADO_EXACTO'] == 1:
        motivos.append(
            'DUPLICADO EXACTO'
        )

    if not motivos:
        return 'SIN OBSERVACIONES'

    return '; '.join(motivos)


df_filtrado['MOTIVO_ESTADO'] = (
    df_filtrado.apply(
        obtener_motivo_estado,
        axis=1
    )
)


# ==========================================================
# 22. SEPARAR RESULTADOS
# ==========================================================

df_validos = df_filtrado[
    df_filtrado[
        'ESTADO_REGISTRO'
    ] == 'VALIDO'
].copy()

df_por_revisar = df_filtrado[
    df_filtrado[
        'ESTADO_REGISTRO'
    ] == 'REVISAR'
].copy()

df_invalidos = df_filtrado[
    df_filtrado[
        'ESTADO_REGISTRO'
    ] == 'INVALIDO'
].copy()

df_virtualidad_revisar = df_filtrado[
    df_filtrado[
        'FLAG_VIRTUAL_EN_CAMPUS_FISICO'
    ] == 1
].copy()

df_dia_revisar = df_filtrado[
    (
        df_filtrado[
            'FLAG_DIA_INCONSISTENTE'
        ] == 1
    )
    |
    (
        df_filtrado[
            'FLAG_DIA_FALTANTE'
        ] == 1
    )
].copy()

df_horas_iguales = df_filtrado[
    df_filtrado[
        'FLAG_HORAS_IGUALES'
    ] == 1
].copy()

df_ceros_registrados = df_filtrado[
    df_filtrado[
        'FLAG_REGISTRADOS_CERO'
    ] == 1
].copy()

df_eventos_repetidos = df_filtrado[
    (
        df_filtrado[
            'FLAG_EVENTO_REPETIDO'
        ] == 1
    )
    |
    (
        df_filtrado[
            'FLAG_DUPLICADO_EXACTO'
        ] == 1
    )
].copy()


# ==========================================================
# 23. CREAR RESUMEN
# ==========================================================

resumen = pd.DataFrame({
    'METRICA': [
        'REGISTROS ORIGINALES',
        'REGISTROS DEL PERIODO Y CAMPUS',
        'REGISTROS VALIDOS',
        'REGISTROS POR REVISAR',
        'REGISTROS INVALIDOS',
        'FECHAS INVALIDAS',
        'DIAS INCONSISTENTES',
        'DIAS FALTANTES',
        'TIPOS DE EXAMEN INVALIDOS',
        'HORARIOS INVALIDOS',
        'HORAS DE INICIO Y FIN IGUALES',
        'DURACIONES ATIPICAS',
        'VIRTUALIDAD EN CAMPUS FISICO',
        'ESVIRTUAL NO RESUELTO',
        'REGISTRADOS EN CERO',
        'REGISTRADOS FALTANTES',
        'REGISTRADOS NEGATIVOS',
        'UBICACIONES INVALIDAS',
        'EVENTOS REPETIDOS',
        'DUPLICADOS EXACTOS',
        'EVENTOS FISICOS PROVISIONALES'
    ],

    'CANTIDAD': [
        cantidad_original,
        len(df_filtrado),
        len(df_validos),
        len(df_por_revisar),
        len(df_invalidos),

        int(
            df_filtrado[
                'FLAG_FECHA_INVALIDA'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_DIA_INCONSISTENTE'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_DIA_FALTANTE'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_TIPO_EXAMEN_INVALIDO'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_HORA_INVALIDA'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_HORAS_IGUALES'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_DURACION_ATIPICA'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_VIRTUAL_EN_CAMPUS_FISICO'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_ESVIRTUAL_NO_RESUELTO'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_REGISTRADOS_CERO'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_REGISTRADOS_FALTANTE'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_REGISTRADOS_NEGATIVO'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_UBICACION_INVALIDA'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_EVENTO_REPETIDO'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_DUPLICADO_EXACTO'
            ].sum()
        ),

        int(
            df_filtrado[
                'ES_EVENTO_FISICO'
            ].sum()
        )
    ]
})


# ==========================================================
# 24. EXPORTAR RESULTADOS
# ==========================================================

os.makedirs(
    os.path.dirname(ruta_salida),
    exist_ok=True
)

with pd.ExcelWriter(
    ruta_salida,
    engine='openpyxl'
) as writer:

    df_filtrado.to_excel(
        writer,
        sheet_name='TODOS_FILTRADOS',
        index=False
    )

    df_validos.to_excel(
        writer,
        sheet_name='VALIDOS',
        index=False
    )

    df_por_revisar.to_excel(
        writer,
        sheet_name='POR_REVISAR',
        index=False
    )

    df_invalidos.to_excel(
        writer,
        sheet_name='INVALIDOS',
        index=False
    )

    df_virtualidad_revisar.to_excel(
        writer,
        sheet_name='VIRTUALIDAD_REVISAR',
        index=False
    )

    df_dia_revisar.to_excel(
        writer,
        sheet_name='DIA_REVISAR',
        index=False
    )

    df_horas_iguales.to_excel(
        writer,
        sheet_name='HORAS_IGUALES',
        index=False
    )

    df_ceros_registrados.to_excel(
        writer,
        sheet_name='CEROS_REGISTRADOS',
        index=False
    )

    df_eventos_repetidos.to_excel(
        writer,
        sheet_name='EVENTOS_REPETIDOS',
        index=False
    )

    resumen.to_excel(
        writer,
        sheet_name='RESUMEN',
        index=False
    )


# ==========================================================
# 25. MOSTRAR RESULTADOS
# ==========================================================

print()
print('LIMPIEZA DE EXÁMENES FINALIZADA')
print('-------------------------------')

print(
    resumen.to_string(
        index=False
    )
)

print()
print(
    f'Archivo generado en:\n{ruta_salida}'
)