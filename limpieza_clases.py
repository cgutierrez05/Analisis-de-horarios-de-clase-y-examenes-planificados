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
# 17. IDENTIFICAR EVENTOS FÍSICOS
# ==========================================================

condicion_evento_fisico = (
    (
        df_filtrado['MODALIDAD_LIMPIA']
        == 'PRESENCIAL'
    )
    & (df_filtrado['ESVIRTUAL'] == 'N')
    & (
        df_filtrado[
            'FLAG_UBICACION_INVALIDA'
        ] == 0
    )
    & (
        df_filtrado[
            'FLAG_MODALIDAD_INCONSISTENTE'
        ] == 0
    )
    & (
        df_filtrado[
            'FLAG_MODALIDAD_NO_RESUELTA'
        ] == 0
    )
)

df_filtrado['ES_EVENTO_FISICO'] = (
    condicion_evento_fisico
    .fillna(False)
    .astype(int)
)


# ==========================================================
# 18. DETECTAR POSIBLES DUPLICADOS
# ==========================================================

columnas_evento = [
    'DIA',
    'HORAINICIO',
    'HORAFIN',
    'CODIGOMATERIA',
    'PARALELO',
    'BLOQUE',
    'AULA'
]

df_filtrado[
    'CANTIDAD_REPETICIONES'
] = (
    df_filtrado
    .groupby(
        columnas_evento,
        dropna=False
    )['ID_FILA_ORIGINAL']
    .transform('size')
)

df_filtrado[
    'FLAG_POSIBLE_DUPLICADO'
] = (
    df_filtrado[
        'CANTIDAD_REPETICIONES'
    ]
    .gt(1)
    .fillna(False)
    .astype(int)
)


# ==========================================================
# 19. ASIGNAR ESTADO A CADA REGISTRO
# ==========================================================

condicion_excluido = (
    df_filtrado[
        'FLAG_VIRTUAL_CONFIRMADO'
    ] == 1
)

condicion_invalido = (
    ~condicion_excluido
    &
    (
        (
            df_filtrado[
                'FLAG_HORA_INVALIDA'
            ] == 1
        )
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
)

condicion_revisar = (
    ~condicion_excluido
    & ~condicion_invalido
    &
    (
        (
            df_filtrado[
                'FLAG_MODALIDAD_INCONSISTENTE'
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
        |
        (
            df_filtrado[
                'FLAG_REGISTRADOS_CERO_PENDIENTE'
            ] == 1
        )
        |
        (
            df_filtrado[
                'FLAG_PRACTICO_TEORICO_AMBIGUO'
            ] == 1
        )
        |
        (
            df_filtrado[
                'FLAG_PRACTICO_SIN_TEORICO'
            ] == 1
        )
        |
        (
            df_filtrado[
                'FLAG_PRACTICO_SIN_REGLA'
            ] == 1
        )
        |
        (
            df_filtrado[
                'FLAG_POSIBLE_DUPLICADO'
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

df_filtrado.loc[
    condicion_excluido,
    'ESTADO_REGISTRO'
] = 'EXCLUIDO_VIRTUAL'


# ==========================================================
# 20. AGREGAR MOTIVO DEL ESTADO
# ==========================================================

def obtener_motivo_estado(fila):
    motivos = []

    if fila['FLAG_MODALIDAD_INFERIDA'] == 1:
        motivos.append(
            'MODALIDAD INFERIDA COMO PRESENCIAL'
        )

    if fila['FLAG_MODALIDAD_INCONSISTENTE'] == 1:
        motivos.append(
            'MODALIDAD Y ESVIRTUAL SE CONTRADICEN'
        )

    if fila['FLAG_MODALIDAD_NO_RESUELTA'] == 1:
        motivos.append(
            'MODALIDAD NO RESUELTA'
        )

    if fila['FLAG_VIRTUAL_CONFIRMADO'] == 1:
        motivos.append(
            'ACTIVIDAD VIRTUAL CONFIRMADA'
        )

    if fila['FLAG_HORA_INVALIDA'] == 1:
        motivos.append(
            'HORARIO INVALIDO'
        )

    if fila['FLAG_DURACION_ATIPICA'] == 1:
        motivos.append(
            'DURACION MAYOR A 8 HORAS'
        )

    if fila['FLAG_REGISTRADOS_FALTANTE'] == 1:
        motivos.append(
            'NUMERO DE REGISTRADOS FALTANTE'
        )

    if fila['FLAG_REGISTRADOS_NEGATIVO'] == 1:
        motivos.append(
            'NUMERO DE REGISTRADOS NEGATIVO'
        )

    if fila['FLAG_PRACTICO_INDEPENDIENTE'] == 1:
        motivos.append(
            'PRACTICO CON REGISTROS PROPIOS'
        )

    if fila['FLAG_PRACTICO_IMPUTADO'] == 1:
        motivos.append(
            'REGISTRADOS HEREDADOS DEL TEORICO'
        )

    if fila['FLAG_PRACTICO_TEORICO_AMBIGUO'] == 1:
        motivos.append(
            'EL TEORICO TIENE VARIOS VALORES'
        )

    if fila['FLAG_PRACTICO_SIN_TEORICO'] == 1:
        motivos.append(
            'NO SE ENCONTRO EL TEORICO ASOCIADO'
        )

    if fila['FLAG_PRACTICO_SIN_REGLA'] == 1:
        motivos.append(
            'NO SE PUDO CALCULAR EL TEORICO ASOCIADO'
        )

    if fila['FLAG_REGISTRADOS_CERO_PENDIENTE'] == 1:
        motivos.append(
            'NUMERO DE REGISTRADOS CONTINUA EN CERO'
        )

    if fila['FLAG_UBICACION_INVALIDA'] == 1:
        motivos.append(
            'AULA O BLOQUE INVALIDO'
        )

    if fila['FLAG_POSIBLE_DUPLICADO'] == 1:
        motivos.append(
            'POSIBLE EVENTO DUPLICADO'
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
# 21. SEPARAR RESULTADOS
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

df_excluidos = df_filtrado[
    df_filtrado[
        'ESTADO_REGISTRO'
    ] == 'EXCLUIDO_VIRTUAL'
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

df_modalidad_revisar = df_filtrado[
    df_filtrado[
        'FLAG_MODALIDAD_INCONSISTENTE'
    ] == 1
].copy()


# ==========================================================
# 22. COMPROBACIONES
# ==========================================================

# Un práctico independiente nunca debe estar asociado
# a un paralelo teórico.
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

# Ningún práctico independiente debe haber sido modificado.
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
        'REGISTROS VIRTUALES EXCLUIDOS',
        'MODALIDADES INFERIDAS',
        'MODALIDADES INCONSISTENTES',
        'REGISTRADOS EN CERO ORIGINAL',
        'PRACTICOS INDEPENDIENTES',
        'PRACTICOS IMPUTADOS',
        'CEROS TODAVIA PENDIENTES',
        'PRACTICOS CON TEORICO AMBIGUO',
        'PRACTICOS SIN TEORICO',
        'PRACTICOS SIN REGLA',
        'POSIBLES DUPLICADOS',
        'INDEPENDIENTES MAL ASOCIADOS',
        'INDEPENDIENTES MODIFICADOS'
    ],
    'CANTIDAD': [
        cantidad_original,
        len(df_filtrado),
        len(df_validos),
        len(df_por_revisar),
        len(df_invalidos),
        len(df_excluidos),
        int(
            df_filtrado[
                'FLAG_MODALIDAD_INFERIDA'
            ].sum()
        ),
        int(
            df_filtrado[
                'FLAG_MODALIDAD_INCONSISTENTE'
            ].sum()
        ),
        int(
            df_filtrado[
                'FLAG_REGISTRADOS_CERO_ORIGINAL'
            ].sum()
        ),
        int(
            df_filtrado[
                'FLAG_PRACTICO_INDEPENDIENTE'
            ].sum()
        ),
        int(
            df_filtrado[
                'FLAG_PRACTICO_IMPUTADO'
            ].sum()
        ),
        int(
            df_filtrado[
                'FLAG_REGISTRADOS_CERO_PENDIENTE'
            ].sum()
        ),
        int(
            df_filtrado[
                'FLAG_PRACTICO_TEORICO_AMBIGUO'
            ].sum()
        ),
        int(
            df_filtrado[
                'FLAG_PRACTICO_SIN_TEORICO'
            ].sum()
        ),
        int(
            df_filtrado[
                'FLAG_PRACTICO_SIN_REGLA'
            ].sum()
        ),
        int(
            df_filtrado[
                'FLAG_POSIBLE_DUPLICADO'
            ].sum()
        ),
        len(
            practicos_independientes_mal_asociados
        ),
        len(
            practicos_independientes_modificados
        )
    ]
})


# ==========================================================
# 24. CREAR CATÁLOGO DE BLOQUES
# ==========================================================

catalogo_bloques = (
    df_filtrado[
        [
            'BLOQUE',
            'REFERENCIA'
        ]
    ]
    .drop_duplicates()
    .sort_values(
        [
            'BLOQUE',
            'REFERENCIA'
        ]
    )
    .reset_index(drop=True)
)


# ==========================================================
# 25. EXPORTAR RESULTADOS
# ==========================================================

os.makedirs(
    os.path.dirname(ruta_salida),
    exist_ok=True
)

os.makedirs(
    os.path.dirname(ruta_catalogo_bloques),
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

    df_excluidos.to_excel(
        writer,
        sheet_name='EXCLUIDOS',
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

    df_modalidad_revisar.to_excel(
        writer,
        sheet_name='MODALIDAD_REVISAR',
        index=False
    )

    resumen.to_excel(
        writer,
        sheet_name='RESUMEN',
        index=False
    )

catalogo_bloques.to_excel(
    ruta_catalogo_bloques,
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
    'Valores NA en condicion_practico_imputable:',
    int(
        condicion_practico_imputable
        .isna()
        .sum()
    )
)

print(
    'Prácticos independientes asociados '
    'incorrectamente:',
    len(
        practicos_independientes_mal_asociados
    )
)

print(
    'Prácticos independientes modificados:',
    len(
        practicos_independientes_modificados
    )
)

print()
print(
    f'Archivo de resultados:\n{ruta_salida}'
)

print()
print(
    f'Catálogo de bloques:\n'
    f'{ruta_catalogo_bloques}'
)