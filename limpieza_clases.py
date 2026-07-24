import os

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
    'horario_clases_limpieza.xlsx'
)

ruta_salida = os.path.join(
    ruta_proyecto,
    'datos_limpios',
    'resultado_limpieza_clases.xlsx'
)

ruta_catalogo_bloques = os.path.join(
    ruta_proyecto,
    'catalogos',
    'catalogo_bloques.xlsx'
)

anio_objetivo = 2026
termino_objetivo = '1S'
campus_objetivo = 'CAMPUS GUSTAVO GALINDO'

# Una duración mayor a ocho horas se marcará como atípica.
duracion_maxima_minutos = 480


# ==========================================================
# 2. CARGAR EL ARCHIVO
# ==========================================================

if not os.path.exists(ruta_entrada):
    raise FileNotFoundError(
        f'No se encontró el archivo:\n{ruta_entrada}'
    )

df = pd.read_excel(ruta_entrada)

if df.empty:
    raise ValueError(
        'El archivo de clases está vacío.'
    )

# La primera fila del Excel contiene los encabezados.
df['ID_FILA_ORIGINAL'] = df.index + 2

cantidad_original = len(df)


# ==========================================================
# 3. LIMPIAR COLUMNAS DE TEXTO
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
# 4. CONVERTIR HORAS
# ==========================================================

hora_inicio_datetime = pd.to_datetime(
    df['HORAINICIO'],
    errors='coerce'
)

hora_fin_datetime = pd.to_datetime(
    df['HORAFIN'],
    errors='coerce'
)

df['DURACIONMINUTOS'] = (
    hora_fin_datetime
    - hora_inicio_datetime
).dt.total_seconds() / 60

# Conservamos únicamente la hora para visualizarla.
df['HORAINICIO'] = hora_inicio_datetime.dt.time
df['HORAFIN'] = hora_fin_datetime.dt.time


# ==========================================================
# 5. CONVERTIR COLUMNAS NUMÉRICAS
# ==========================================================

df['ANIO'] = pd.to_numeric(
    df['ANIO'],
    errors='coerce'
).astype('Int64')

df['PARALELO'] = pd.to_numeric(
    df['PARALELO'],
    errors='coerce'
).astype('Int64')

# Conservamos el dato original antes de transformarlo.
df['NUMREGISTRADOS_REPORTADO'] = (
    df['NUMREGISTRADOS']
)

df['NUMREGISTRADOS'] = pd.to_numeric(
    df['NUMREGISTRADOS'],
    errors='coerce'
)


# ==========================================================
# 6. FILTRAR PERIODO Y CAMPUS
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
# 7. LIMPIAR E INFERIR MODALIDAD
# ==========================================================

df_filtrado['MODALIDAD_REPORTADA'] = (
    df_filtrado['MODALIDAD']
)

df_filtrado['MODALIDAD_LIMPIA'] = (
    df_filtrado['MODALIDAD']
)

# Si la modalidad está vacía, pero el evento:
# - No está marcado como virtual.
# - Tiene aula.
# - Tiene bloque.
# Se infiere provisionalmente que es presencial.
condicion_modalidad_inferida = (
    df_filtrado['MODALIDAD'].isna()
    & (df_filtrado['ESVIRTUAL'] == 'N')
    & df_filtrado['AULA'].notna()
    & df_filtrado['BLOQUE'].notna()
)

df_filtrado.loc[
    condicion_modalidad_inferida,
    'MODALIDAD_LIMPIA'
] = 'PRESENCIAL'

df_filtrado['FLAG_MODALIDAD_INFERIDA'] = (
    condicion_modalidad_inferida
    .fillna(False)
    .astype(int)
)

# Detectamos contradicciones.
condicion_modalidad_inconsistente = (
    (
        (
            df_filtrado['MODALIDAD_LIMPIA']
            == 'PRESENCIAL'
        )
        & (df_filtrado['ESVIRTUAL'] == 'S')
    )
    |
    (
        (
            df_filtrado['MODALIDAD_LIMPIA']
            == 'VIRTUAL'
        )
        & (df_filtrado['ESVIRTUAL'] == 'N')
    )
).fillna(False)

df_filtrado[
    'FLAG_MODALIDAD_INCONSISTENTE'
] = condicion_modalidad_inconsistente.astype(int)

# Modalidades que no pudieron interpretarse.
condicion_modalidad_no_resuelta = (
    df_filtrado['MODALIDAD_LIMPIA'].isna()
    |
    ~df_filtrado['MODALIDAD_LIMPIA'].isin(
        ['PRESENCIAL', 'VIRTUAL']
    )
    |
    ~df_filtrado['ESVIRTUAL'].isin(
        ['S', 'N']
    )
)

df_filtrado[
    'FLAG_MODALIDAD_NO_RESUELTA'
] = condicion_modalidad_no_resuelta.astype(int)

# Evento virtual confirmado:
# ambas columnas coinciden en que es virtual.
condicion_virtual_confirmada = (
    (
        df_filtrado['MODALIDAD_LIMPIA']
        == 'VIRTUAL'
    )
    & (df_filtrado['ESVIRTUAL'] == 'S')
).fillna(False)

df_filtrado[
    'FLAG_VIRTUAL_CONFIRMADO'
] = condicion_virtual_confirmada.astype(int)


# ==========================================================
# 8. VALIDAR HORARIOS
# ==========================================================

condicion_hora_invalida = (
    df_filtrado['HORAINICIO'].isna()
    | df_filtrado['HORAFIN'].isna()
    | df_filtrado['DURACIONMINUTOS'].isna()
    | (df_filtrado['DURACIONMINUTOS'] <= 0)
)

df_filtrado['FLAG_HORA_INVALIDA'] = (
    condicion_hora_invalida
    .fillna(False)
    .astype(int)
)

condicion_duracion_atipica = (
    (df_filtrado['FLAG_HORA_INVALIDA'] == 0)
    & (
        df_filtrado['DURACIONMINUTOS']
        > duracion_maxima_minutos
    )
).fillna(False)

df_filtrado['FLAG_DURACION_ATIPICA'] = (
    condicion_duracion_atipica.astype(int)
)


# ==========================================================
# 9. VALIDAR NÚMERO DE ESTUDIANTES
# ==========================================================

df_filtrado['FLAG_REGISTRADOS_FALTANTE'] = (
    df_filtrado['NUMREGISTRADOS']
    .isna()
    .astype(int)
)

df_filtrado['FLAG_REGISTRADOS_NEGATIVO'] = (
    df_filtrado['NUMREGISTRADOS']
    .lt(0)
    .fillna(False)
    .astype(int)
)

df_filtrado[
    'FLAG_REGISTRADOS_CERO_ORIGINAL'
] = (
    df_filtrado['NUMREGISTRADOS']
    .eq(0)
    .fillna(False)
    .astype(int)
)

# Inicialmente, el valor ajustado es igual al original.
df_filtrado['NUMREGISTRADOS_AJUSTADO'] = (
    df_filtrado['NUMREGISTRADOS']
)


# ==========================================================
# 10. VALIDAR UBICACIÓN
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

df_filtrado['FLAG_UBICACION_INVALIDA'] = (
    condicion_ubicacion_invalida
    .fillna(False)
    .astype(int)
)

# Una misma aula puede aparecer en distintos bloques.
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
# 11. IDENTIFICAR PARALELOS TEÓRICOS Y PRÁCTICOS
# ==========================================================

condicion_paralelo_teorico = (
    df_filtrado['PARALELO']
    .lt(100)
    .fillna(False)
)

condicion_paralelo_practico = (
    df_filtrado['PARALELO']
    .ge(100)
    .fillna(False)
)

# Prácticos con estudiantes propios:
# no deben relacionarse con ningún teórico.
condicion_practico_con_registros = (
    condicion_paralelo_practico
    & df_filtrado['NUMREGISTRADOS']
    .gt(0)
    .fillna(False)
)

# Solamente los prácticos con cero buscarán
# un paralelo teórico asociado.
condicion_practico_cero = (
    condicion_paralelo_practico
    & df_filtrado['NUMREGISTRADOS']
    .eq(0)
    .fillna(False)
)

df_filtrado['TIPO_PARALELO'] = np.select(
    [
        condicion_paralelo_teorico,
        condicion_practico_con_registros,
        condicion_practico_cero
    ],
    [
        'TEORICO',
        'PRACTICO_INDEPENDIENTE',
        'PRACTICO_SIN_REGISTRADOS'
    ],
    default='DESCONOCIDO'
)

# Se deja vacío para todos los registros inicialmente.
df_filtrado[
    'PARALELO_TEORICO_ASOCIADO'
] = pd.Series(
    pd.NA,
    index=df_filtrado.index,
    dtype='Int64'
)

paralelo_teorico_calculado = (
    df_filtrado['PARALELO'] % 100
)

# Solo se asigna un teórico a los prácticos
# cuyo número de registrados es cero.
condicion_asociacion_valida = (
    condicion_practico_cero
    & paralelo_teorico_calculado
    .gt(0)
    .fillna(False)
)

df_filtrado.loc[
    condicion_asociacion_valida,
    'PARALELO_TEORICO_ASOCIADO'
] = paralelo_teorico_calculado[
    condicion_asociacion_valida
].astype('Int64')


# ==========================================================
# 12. CONSTRUIR TABLA DE PARALELOS TEÓRICOS
# ==========================================================

tabla_teoricos = (
    df_filtrado[
        (
            df_filtrado['TIPO_PARALELO']
            == 'TEORICO'
        )
        & df_filtrado['NUMREGISTRADOS']
        .gt(0)
        .fillna(False)
    ]
    .groupby(
        [
            'CODIGOMATERIA',
            'PARALELO'
        ],
        as_index=False
    )
    .agg(
        NUMREGISTRADOS_TEORICO=(
            'NUMREGISTRADOS',
            'first'
        ),
        CANTIDAD_VALORES_TEORICO=(
            'NUMREGISTRADOS',
            'nunique'
        )
    )
)

tabla_teoricos = tabla_teoricos.rename(
    columns={
        'PARALELO':
        'PARALELO_TEORICO_ASOCIADO'
    }
)

tabla_teoricos[
    'PARALELO_TEORICO_ASOCIADO'
] = tabla_teoricos[
    'PARALELO_TEORICO_ASOCIADO'
].astype('Int64')


# ==========================================================
# 13. UNIR PRÁCTICOS CON TEÓRICOS
# ==========================================================

df_filtrado = df_filtrado.merge(
    tabla_teoricos,
    on=[
        'CODIGOMATERIA',
        'PARALELO_TEORICO_ASOCIADO'
    ],
    how='left',
    validate='many_to_one'
)

# El merge reconstruye el índice.
df_filtrado = df_filtrado.reset_index(
    drop=True
)


# ==========================================================
# 14. RECALCULAR CONDICIONES DESPUÉS DEL MERGE
# ==========================================================

# Es necesario recalcular estas condiciones.
# Las condiciones anteriores conservaban el índice
# existente antes del merge.

condicion_paralelo_practico = (
    df_filtrado['PARALELO']
    .ge(100)
    .fillna(False)
)

condicion_practico_con_registros = (
    condicion_paralelo_practico
    & df_filtrado['NUMREGISTRADOS']
    .gt(0)
    .fillna(False)
)

condicion_practico_cero = (
    condicion_paralelo_practico
    & df_filtrado['NUMREGISTRADOS']
    .eq(0)
    .fillna(False)
)


# ==========================================================
# 15. IMPUTAR PARALELOS PRÁCTICOS
# ==========================================================

condicion_practico_imputable = (
    condicion_practico_cero
    & df_filtrado[
        'PARALELO_TEORICO_ASOCIADO'
    ].notna()
    & df_filtrado[
        'NUMREGISTRADOS_TEORICO'
    ].notna()
    & df_filtrado[
        'CANTIDAD_VALORES_TEORICO'
    ].eq(1)
).fillna(False)

df_filtrado.loc[
    condicion_practico_imputable,
    'NUMREGISTRADOS_AJUSTADO'
] = df_filtrado.loc[
    condicion_practico_imputable,
    'NUMREGISTRADOS_TEORICO'
]

df_filtrado['FLAG_PRACTICO_IMPUTADO'] = (
    condicion_practico_imputable
    .fillna(False)
    .astype(int)
)

df_filtrado[
    'FLAG_PRACTICO_INDEPENDIENTE'
] = (
    condicion_practico_con_registros
    .fillna(False)
    .astype(int)
)

# El teórico asociado presenta más de un
# número de estudiantes.
condicion_teorico_ambiguo = (
    condicion_practico_cero
    & df_filtrado[
        'CANTIDAD_VALORES_TEORICO'
    ].gt(1)
).fillna(False)

df_filtrado[
    'FLAG_PRACTICO_TEORICO_AMBIGUO'
] = condicion_teorico_ambiguo.astype(int)

# Se calculó el teórico asociado, pero no se encontró
# un valor positivo para ese teórico.
condicion_practico_sin_teorico = (
    condicion_practico_cero
    & df_filtrado[
        'PARALELO_TEORICO_ASOCIADO'
    ].notna()
    & df_filtrado[
        'NUMREGISTRADOS_TEORICO'
    ].isna()
).fillna(False)

df_filtrado[
    'FLAG_PRACTICO_SIN_TEORICO'
] = condicion_practico_sin_teorico.astype(int)

# Casos como 100, 200 o 300 cuyo residuo es cero.
condicion_practico_sin_regla = (
    condicion_practico_cero
    & df_filtrado[
        'PARALELO_TEORICO_ASOCIADO'
    ].isna()
).fillna(False)

df_filtrado[
    'FLAG_PRACTICO_SIN_REGLA'
] = condicion_practico_sin_regla.astype(int)


# ==========================================================
# 16. REGISTRAR ORIGEN DEL NÚMERO DE ESTUDIANTES
# ==========================================================

df_filtrado[
    'ORIGEN_NUMREGISTRADOS'
] = 'ORIGINAL'

df_filtrado.loc[
    condicion_practico_con_registros,
    'ORIGEN_NUMREGISTRADOS'
] = 'ORIGINAL_PRACTICO_INDEPENDIENTE'

df_filtrado.loc[
    condicion_practico_imputable,
    'ORIGEN_NUMREGISTRADOS'
] = 'HEREDADO_DEL_PARALELO_TEORICO'

df_filtrado.loc[
    condicion_teorico_ambiguo,
    'ORIGEN_NUMREGISTRADOS'
] = 'TEORICO_AMBIGUO'

df_filtrado.loc[
    condicion_practico_sin_teorico,
    'ORIGEN_NUMREGISTRADOS'
] = 'TEORICO_NO_ENCONTRADO'

df_filtrado.loc[
    condicion_practico_sin_regla,
    'ORIGEN_NUMREGISTRADOS'
] = 'ASOCIACION_NO_DEFINIDA'

df_filtrado[
    'NUMREGISTRADOS_AJUSTADO'
] = (
    pd.to_numeric(
        df_filtrado[
            'NUMREGISTRADOS_AJUSTADO'
        ],
        errors='coerce'
    )
    .round()
    .astype('Int64')
)

# Ceros que continúan pendientes después
# de intentar completar los prácticos.
df_filtrado[
    'FLAG_REGISTRADOS_CERO_PENDIENTE'
] = (
    df_filtrado[
        'NUMREGISTRADOS_AJUSTADO'
    ]
    .eq(0)
    .fillna(False)
    .astype(int)
)


# ==========================================================
# 17. RESOLVER INCONSISTENCIAS DE MODALIDAD
# ==========================================================

# Conservamos la virtualidad original.
df_filtrado['ESVIRTUAL_REPORTADO'] = (
    df_filtrado['ESVIRTUAL']
)

# Creamos una versión limpia que podrá corregirse.
df_filtrado['ESVIRTUAL_LIMPIO'] = (
    df_filtrado['ESVIRTUAL']
)

# Conservamos también la bandera original de inconsistencia.
df_filtrado[
    'FLAG_MODALIDAD_INCONSISTENTE_ORIGINAL'
] = df_filtrado[
    'FLAG_MODALIDAD_INCONSISTENTE'
]

# Una inconsistencia se resuelve como presencial cuando:
# - Tiene estudiantes después de la imputación.
# - Tiene aula.
# - Tiene bloque.
# - La ubicación no está marcada como inválida.
condicion_modalidad_corregida_presencial = (
    (
        df_filtrado[
            'FLAG_MODALIDAD_INCONSISTENTE_ORIGINAL'
        ] == 1
    )
    & df_filtrado[
        'NUMREGISTRADOS_AJUSTADO'
    ].gt(0).fillna(False)
    & df_filtrado['AULA'].notna()
    & df_filtrado['BLOQUE'].notna()
    & (
        df_filtrado[
            'FLAG_UBICACION_INVALIDA'
        ] == 0
    )
)

df_filtrado[
    'FLAG_MODALIDAD_CORREGIDA_PRESENCIAL'
] = (
    condicion_modalidad_corregida_presencial
    .fillna(False)
    .astype(int)
)

df_filtrado.loc[
    condicion_modalidad_corregida_presencial,
    'MODALIDAD_LIMPIA'
] = 'PRESENCIAL'

df_filtrado.loc[
    condicion_modalidad_corregida_presencial,
    'ESVIRTUAL_LIMPIO'
] = 'N'

# Las inconsistencias que no pudieron corregirse
# permanecerán pendientes.
df_filtrado[
    'FLAG_MODALIDAD_INCONSISTENTE_PENDIENTE'
] = (
    (
        df_filtrado[
            'FLAG_MODALIDAD_INCONSISTENTE_ORIGINAL'
        ] == 1
    )
    & ~condicion_modalidad_corregida_presencial
).fillna(False).astype(int)

# Recalculamos la modalidad no resuelta con las columnas limpias.
condicion_modalidad_no_resuelta_final = (
    df_filtrado['MODALIDAD_LIMPIA'].isna()
    |
    ~df_filtrado['MODALIDAD_LIMPIA'].isin(
        ['PRESENCIAL', 'VIRTUAL']
    )
    |
    df_filtrado['ESVIRTUAL_LIMPIO'].isna()
    |
    ~df_filtrado['ESVIRTUAL_LIMPIO'].isin(
        ['S', 'N']
    )
)

df_filtrado[
    'FLAG_MODALIDAD_NO_RESUELTA'
] = (
    condicion_modalidad_no_resuelta_final
    .fillna(False)
    .astype(int)
)

# Recalculamos la virtualidad confirmada.
condicion_virtual_confirmada = (
    (
        df_filtrado['MODALIDAD_LIMPIA']
        == 'VIRTUAL'
    )
    & (
        df_filtrado['ESVIRTUAL_LIMPIO']
        == 'S'
    )
)

df_filtrado[
    'FLAG_VIRTUAL_CONFIRMADO'
] = (
    condicion_virtual_confirmada
    .fillna(False)
    .astype(int)
)


# ==========================================================
# 18. IDENTIFICAR REGISTROS CON CERO ESTUDIANTES
# ==========================================================

condicion_excluir_cero = (
    df_filtrado[
        'NUMREGISTRADOS_AJUSTADO'
    ]
    .eq(0)
    .fillna(False)
)

df_filtrado[
    'FLAG_EXCLUIR_CERO'
] = condicion_excluir_cero.astype(int)


# ==========================================================
# 19. IDENTIFICAR EVENTOS FÍSICOS
# ==========================================================

condicion_evento_fisico = (
    (
        df_filtrado['MODALIDAD_LIMPIA']
        == 'PRESENCIAL'
    )
    & (
        df_filtrado['ESVIRTUAL_LIMPIO']
        == 'N'
    )
    & (
        df_filtrado[
            'FLAG_UBICACION_INVALIDA'
        ] == 0
    )
    & (
        df_filtrado[
            'FLAG_MODALIDAD_INCONSISTENTE_PENDIENTE'
        ] == 0
    )
    & (
        df_filtrado[
            'FLAG_MODALIDAD_NO_RESUELTA'
        ] == 0
    )
    & df_filtrado[
        'NUMREGISTRADOS_AJUSTADO'
    ].gt(0).fillna(False)
)

df_filtrado['ES_EVENTO_FISICO'] = (
    condicion_evento_fisico
    .fillna(False)
    .astype(int)
)


# ==========================================================
# 20. DETECTAR Y ELIMINAR DUPLICADOS
# ==========================================================

# Se conserva la misma definición de evento
# usada en el script anterior.
columnas_evento = [
    'DIA',
    'HORAINICIO',
    'HORAFIN',
    'CODIGOMATERIA',
    'PARALELO',
    'BLOQUE',
    'AULA'
]

df_filtrado = (
    df_filtrado
    .sort_values('ID_FILA_ORIGINAL')
    .reset_index(drop=True)
)

condicion_candidata_duplicado = (
    df_filtrado['ES_EVENTO_FISICO'] == 1
)

df_filtrado[
    'CANTIDAD_REPETICIONES'
] = pd.Series(
    0,
    index=df_filtrado.index,
    dtype='Int64'
)

conteo_repeticiones = (
    df_filtrado.loc[
        condicion_candidata_duplicado
    ]
    .groupby(
        columnas_evento,
        dropna=False
    )['ID_FILA_ORIGINAL']
    .transform('size')
)

df_filtrado.loc[
    condicion_candidata_duplicado,
    'CANTIDAD_REPETICIONES'
] = conteo_repeticiones.astype('Int64')

# Esta bandera marca todas las filas que pertenecen a un grupo repetido, incluida la primera copia.
df_filtrado[
    'FLAG_GRUPO_DUPLICADO'
] = 0

df_filtrado.loc[
    condicion_candidata_duplicado,
    'FLAG_GRUPO_DUPLICADO'
] = (
    df_filtrado.loc[
        condicion_candidata_duplicado
    ]
    .duplicated(
        subset=columnas_evento,
        keep=False
    )
    .astype(int)
)

# Esta bandera marca solamente las copias que deben eliminarse. La primera fila se conserva.
df_filtrado[
    'FLAG_DUPLICADO_ELIMINAR'
] = 0

df_filtrado.loc[
    condicion_candidata_duplicado,
    'FLAG_DUPLICADO_ELIMINAR'
] = (
    df_filtrado.loc[
        condicion_candidata_duplicado
    ]
    .duplicated(
        subset=columnas_evento,
        keep='first'
    )
    .astype(int)
)

# Se conserva esta columna por compatibilidad
# con el resumen anterior.
df_filtrado[
    'FLAG_POSIBLE_DUPLICADO'
] = df_filtrado[
    'FLAG_GRUPO_DUPLICADO'
]


# ==========================================================
# 21. ASIGNAR ESTADO FINAL
# ==========================================================

condicion_invalido = (
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
        (
            df_filtrado[
                'MODALIDAD_LIMPIA'
            ] == 'PRESENCIAL'
        )
        &
        (
            df_filtrado[
                'FLAG_UBICACION_INVALIDA'
            ] == 1
        )
    )
)

condicion_excluido_virtual = (
    ~condicion_invalido
    & (
        df_filtrado[
            'FLAG_VIRTUAL_CONFIRMADO'
        ] == 1
    )
)

condicion_excluido_cero = (
    ~condicion_invalido
    & ~condicion_excluido_virtual
    & (
        df_filtrado[
            'FLAG_EXCLUIR_CERO'
        ] == 1
    )
)

condicion_excluido_duplicado = (
    ~condicion_invalido
    & ~condicion_excluido_virtual
    & ~condicion_excluido_cero
    & (
        df_filtrado[
            'FLAG_DUPLICADO_ELIMINAR'
        ] == 1
    )
)

condicion_revisar = (
    ~condicion_invalido
    & ~condicion_excluido_virtual
    & ~condicion_excluido_cero
    & ~condicion_excluido_duplicado
    &
    (
        (
            df_filtrado[
                'FLAG_MODALIDAD_INCONSISTENTE_PENDIENTE'
            ] == 1
        )
        |
        (
            df_filtrado[
                'FLAG_MODALIDAD_NO_RESUELTA'
            ] == 1
        )
        |
        (
            df_filtrado[
                'FLAG_DURACION_ATIPICA'
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
    condicion_excluido_duplicado,
    'ESTADO_REGISTRO'
] = 'EXCLUIDO_DUPLICADO'

df_filtrado.loc[
    condicion_excluido_cero,
    'ESTADO_REGISTRO'
] = 'EXCLUIDO_CERO'

df_filtrado.loc[
    condicion_excluido_virtual,
    'ESTADO_REGISTRO'
] = 'EXCLUIDO_VIRTUAL'

df_filtrado.loc[
    condicion_invalido,
    'ESTADO_REGISTRO'
] = 'INVALIDO'


# ==========================================================
# 22. CREAR MOTIVO DEL ESTADO
# ==========================================================

def obtener_motivo_estado(fila):

    motivos = []

    if fila[
        'FLAG_MODALIDAD_CORREGIDA_PRESENCIAL'
    ] == 1:
        motivos.append(
            'INCONSISTENCIA RESUELTA COMO PRESENCIAL'
        )

    if fila[
        'FLAG_MODALIDAD_INCONSISTENTE_PENDIENTE'
    ] == 1:
        motivos.append(
            'INCONSISTENCIA DE MODALIDAD PENDIENTE'
        )

    if fila[
        'FLAG_MODALIDAD_NO_RESUELTA'
    ] == 1:
        motivos.append(
            'MODALIDAD NO RESUELTA'
        )

    if fila[
        'FLAG_VIRTUAL_CONFIRMADO'
    ] == 1:
        motivos.append(
            'ACTIVIDAD VIRTUAL CONFIRMADA'
        )

    if fila[
        'FLAG_HORA_INVALIDA'
    ] == 1:
        motivos.append(
            'HORARIO INVALIDO'
        )

    if fila[
        'FLAG_DURACION_ATIPICA'
    ] == 1:
        motivos.append(
            'DURACION MAYOR A 8 HORAS'
        )

    if fila[
        'FLAG_REGISTRADOS_FALTANTE'
    ] == 1:
        motivos.append(
            'NUMERO DE REGISTRADOS FALTANTE'
        )

    if fila[
        'FLAG_REGISTRADOS_NEGATIVO'
    ] == 1:
        motivos.append(
            'NUMERO DE REGISTRADOS NEGATIVO'
        )

    if fila[
        'FLAG_PRACTICO_IMPUTADO'
    ] == 1:
        motivos.append(
            'REGISTRADOS HEREDADOS DEL TEORICO'
        )

    if fila[
        'FLAG_EXCLUIR_CERO'
    ] == 1:
        motivos.append(
            'REGISTRO EXCLUIDO POR CERO ESTUDIANTES'
        )

    if fila[
        'FLAG_DUPLICADO_ELIMINAR'
    ] == 1:
        motivos.append(
            'COPIA DUPLICADA ELIMINADA'
        )

    if (
        fila['FLAG_GRUPO_DUPLICADO'] == 1
        and fila['FLAG_DUPLICADO_ELIMINAR'] == 0
    ):
        motivos.append(
            'PRIMERA COPIA DEL EVENTO CONSERVADA'
        )

    if fila[
        'FLAG_UBICACION_INVALIDA'
    ] == 1:
        motivos.append(
            'AULA O BLOQUE INVALIDO'
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
# 23. SEPARAR RESULTADOS
# ==========================================================

df_limpio_final = df_filtrado[
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

df_excluidos_virtuales = df_filtrado[
    df_filtrado[
        'ESTADO_REGISTRO'
    ] == 'EXCLUIDO_VIRTUAL'
].copy()

df_excluidos_cero = df_filtrado[
    df_filtrado[
        'ESTADO_REGISTRO'
    ] == 'EXCLUIDO_CERO'
].copy()

df_duplicados_eliminados = df_filtrado[
    df_filtrado[
        'ESTADO_REGISTRO'
    ] == 'EXCLUIDO_DUPLICADO'
].copy()

df_modalidad_corregida = df_filtrado[
    df_filtrado[
        'FLAG_MODALIDAD_CORREGIDA_PRESENCIAL'
    ] == 1
].copy()

df_grupos_duplicados = df_filtrado[
    df_filtrado[
        'FLAG_GRUPO_DUPLICADO'
    ] == 1
].copy()

df_practicos_imputados = df_filtrado[
    df_filtrado[
        'FLAG_PRACTICO_IMPUTADO'
    ] == 1
].copy()

df_practicos_independientes = df_filtrado[
    df_filtrado[
        'FLAG_PRACTICO_INDEPENDIENTE'
    ] == 1
].copy()


# ==========================================================
# 24. COMPROBACIONES
# ==========================================================

practicos_independientes_mal_asociados = (
    df_filtrado[
        (
            df_filtrado['TIPO_PARALELO']
            == 'PRACTICO_INDEPENDIENTE'
        )
        & df_filtrado[
            'PARALELO_TEORICO_ASOCIADO'
        ].notna()
    ]
)

practicos_independientes_modificados = (
    df_filtrado[
        (
            df_filtrado['TIPO_PARALELO']
            == 'PRACTICO_INDEPENDIENTE'
        )
        & (
            df_filtrado['NUMREGISTRADOS']
            !=
            df_filtrado[
                'NUMREGISTRADOS_AJUSTADO'
            ]
        )
    ]
)

cantidad_clasificada = (
    len(df_limpio_final)
    + len(df_por_revisar)
    + len(df_invalidos)
    + len(df_excluidos_virtuales)
    + len(df_excluidos_cero)
    + len(df_duplicados_eliminados)
)

if cantidad_clasificada != len(df_filtrado):
    raise ValueError(
        'La clasificación final no coincide con '
        'el número total de registros.'
    )


# ==========================================================
# 25. CREAR RESUMEN
# ==========================================================

resumen = pd.DataFrame({
    'METRICA': [
        'REGISTROS ORIGINALES',
        'REGISTROS DEL PERIODO Y CAMPUS',
        'PRACTICOS IMPUTADOS',
        'MODALIDADES CORREGIDAS COMO PRESENCIAL',
        'REGISTROS EN GRUPOS DUPLICADOS',
        'COPIAS DUPLICADAS ELIMINADAS',
        'REGISTROS EXCLUIDOS POR CERO',
        'REGISTROS VIRTUALES EXCLUIDOS',
        'REGISTROS INVALIDOS',
        'REGISTROS POR REVISAR',
        'REGISTROS LIMPIOS FINALES',
        'INDEPENDIENTES MAL ASOCIADOS',
        'INDEPENDIENTES MODIFICADOS'
    ],

    'CANTIDAD': [
        cantidad_original,
        len(df_filtrado),

        int(
            df_filtrado[
                'FLAG_PRACTICO_IMPUTADO'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_MODALIDAD_CORREGIDA_PRESENCIAL'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_GRUPO_DUPLICADO'
            ].sum()
        ),

        int(
            df_filtrado[
                'FLAG_DUPLICADO_ELIMINAR'
            ].sum()
        ),

        len(df_excluidos_cero),
        len(df_excluidos_virtuales),
        len(df_invalidos),
        len(df_por_revisar),
        len(df_limpio_final),

        len(
            practicos_independientes_mal_asociados
        ),

        len(
            practicos_independientes_modificados
        )
    ]
})


# ==========================================================
# 26. EXPORTAR RESULTADOS
# ==========================================================

os.makedirs(
    os.path.dirname(ruta_salida),
    exist_ok=True
)

with pd.ExcelWriter(
    ruta_salida,
    engine='openpyxl'
) as writer:

    # Contiene todos los registros y todas las banderas.
    df_filtrado.to_excel(
        writer,
        sheet_name='TODOS_FILTRADOS',
        index=False
    )

    # Esta es la hoja que debes utilizar para el análisis.
    df_limpio_final.to_excel(
        writer,
        sheet_name='DATOS_LIMPIOS',
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

    df_excluidos_cero.to_excel(
        writer,
        sheet_name='EXCLUIDOS_CERO',
        index=False
    )

    df_duplicados_eliminados.to_excel(
        writer,
        sheet_name='DUPLICADOS_ELIM',
        index=False
    )

    df_grupos_duplicados.to_excel(
        writer,
        sheet_name='GRUPOS_DUPLICADOS',
        index=False
    )

    df_excluidos_virtuales.to_excel(
        writer,
        sheet_name='EXCLUIDOS_VIRTUAL',
        index=False
    )

    df_modalidad_corregida.to_excel(
        writer,
        sheet_name='MODALIDAD_CORREGIDA',
        index=False
    )

    df_practicos_imputados.to_excel(
        writer,
        sheet_name='PRACTICOS_IMPUTADOS',
        index=False
    )

    df_practicos_independientes.to_excel(
        writer,
        sheet_name='PRACTICOS_INDEPEND',
        index=False
    )

    resumen.to_excel(
        writer,
        sheet_name='RESUMEN',
        index=False
    )

# ==========================================================
# 26. MOSTRAR RESULTADOS
# ==========================================================

print()
print('LIMPIEZA FINALIZADA')
print('-------------------')
print(resumen.to_string(index=False))

print()
print(
    'Registros limpios finales:',
    len(df_limpio_final)
)

print(
    'Registros excluidos por cero:',
    len(df_excluidos_cero)
)

print(
    'Copias duplicadas eliminadas:',
    len(df_duplicados_eliminados)
)

print(
    'Modalidades corregidas como presencial:',
    len(df_modalidad_corregida)
)

print()
print(
    f'Archivo de resultados:\n{ruta_salida}'
)