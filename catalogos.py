import os

import pandas as pd


# ==========================================================
# 1. RUTAS
# ==========================================================

ruta_proyecto = (
    r'/workspaces/'
    r'Analisis-de-horarios-de-clase-y-examenes-planificados'
)

ruta_clases = os.path.join(
    ruta_proyecto,
    'datos_limpios',
    'resultado_limpieza_clases.xlsx'
)

ruta_examenes = os.path.join(
    ruta_proyecto,
    'datos_limpios',
    'resultado_limpieza_examenes.xlsx'
)

ruta_catalogos = os.path.join(
    ruta_proyecto,
    'catalogos'
)

ruta_catalogo_bloques = os.path.join(
    ruta_catalogos,
    'catalogo_bloques.xlsx'
)

os.makedirs(
    ruta_catalogos,
    exist_ok=True
)


# ==========================================================
# 2. FUNCIONES AUXILIARES
# ==========================================================

def limpiar_texto(serie):
    """
    Limpia espacios, convierte a mayúsculas
    y transforma cadenas vacías en valores nulos.
    """

    return (
        serie
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


def unir_valores_unicos(serie):
    """
    Une valores únicos en una sola cadena.

    Ejemplo:
    FIEC | GOBIERNO FIEC
    """

    valores = (
        serie
        .dropna()
        .astype(str)
        .str.strip()
        .tolist()
    )

    valores = sorted(
        set(valores)
    )

    if not valores:
        return pd.NA

    return ' | '.join(valores)


def preparar_ubicaciones(dataframe):
    """
    Extrae solamente BLOQUE y REFERENCIA.

    No utiliza AULA porque la unidad espacial
    del proyecto será el bloque.
    """

    columnas_necesarias = {
        'BLOQUE',
        'REFERENCIA'
    }

    columnas_faltantes = (
        columnas_necesarias
        - set(dataframe.columns)
    )

    if columnas_faltantes:
        raise ValueError(
            'Faltan columnas necesarias: '
            + ', '.join(
                sorted(columnas_faltantes)
            )
        )

    resultado = dataframe[
        [
            'BLOQUE',
            'REFERENCIA'
        ]
    ].copy()

    resultado['BLOQUE'] = limpiar_texto(
        resultado['BLOQUE']
    )

    resultado['REFERENCIA'] = limpiar_texto(
        resultado['REFERENCIA']
    )

    # Eliminamos registros sin bloque.
    resultado = resultado[
        resultado['BLOQUE'].notna()
    ].copy()

    # Eliminamos identificadores que claramente
    # no representan un bloque físico.
    condicion_bloque_virtual = (
        resultado['BLOQUE']
        .fillna('')
        .str.contains(
            'VIRTUAL',
            regex=False
        )
    )

    condicion_sin_bloque = (
        resultado['BLOQUE'].isin(
            [
                'SIN_BLOQUE',
                'SIN BLOQUE',
                'N/A',
                'NA'
            ]
        )
    )

    resultado = resultado[
        ~condicion_bloque_virtual
        & ~condicion_sin_bloque
    ].copy()

    return resultado


def combinar_referencias(fila):
    """
    Combina las referencias encontradas
    en clases y exámenes.
    """

    referencias = []

    columnas_referencia = [
        'REFERENCIAS_CLASES',
        'REFERENCIAS_EXAMENES'
    ]

    for columna in columnas_referencia:

        valor = fila.get(
            columna,
            pd.NA
        )

        if pd.notna(valor):

            partes = str(valor).split('|')

            for parte in partes:

                parte = parte.strip()

                if parte:
                    referencias.append(parte)

    referencias = sorted(
        set(referencias)
    )

    if not referencias:
        return pd.NA

    return ' | '.join(referencias)


# ==========================================================
# 3. COMPROBAR ARCHIVOS
# ==========================================================

if not os.path.exists(ruta_clases):
    raise FileNotFoundError(
        'No se encontró el resultado de clases:\n'
        f'{ruta_clases}'
    )

if not os.path.exists(ruta_examenes):
    raise FileNotFoundError(
        'No se encontró el resultado de exámenes:\n'
        f'{ruta_examenes}'
    )


# ==========================================================
# 4. LEER RESULTADOS DE LIMPIEZA
# ==========================================================

df_clases = pd.read_excel(
    ruta_clases,
    sheet_name='TODOS_FILTRADOS'
)

df_examenes = pd.read_excel(
    ruta_examenes,
    sheet_name='TODOS_FILTRADOS'
)

print(
    f'Registros de clases leídos: '
    f'{len(df_clases)}'
)

print(
    f'Registros de exámenes leídos: '
    f'{len(df_examenes)}'
)


# ==========================================================
# 5. PREPARAR BLOQUES
# ==========================================================

bloques_clases = preparar_ubicaciones(
    df_clases
)

bloques_examenes = preparar_ubicaciones(
    df_examenes
)


# ==========================================================
# 6. RESUMIR BLOQUES DE CLASES
# ==========================================================

resumen_clases = (
    bloques_clases
    .groupby(
        'BLOQUE',
        as_index=False
    )
    .agg(
        REFERENCIAS_CLASES=(
            'REFERENCIA',
            unir_valores_unicos
        ),
        REGISTROS_CLASES=(
            'BLOQUE',
            'size'
        )
    )
)


# ==========================================================
# 7. RESUMIR BLOQUES DE EXÁMENES
# ==========================================================

resumen_examenes = (
    bloques_examenes
    .groupby(
        'BLOQUE',
        as_index=False
    )
    .agg(
        REFERENCIAS_EXAMENES=(
            'REFERENCIA',
            unir_valores_unicos
        ),
        REGISTROS_EXAMENES=(
            'BLOQUE',
            'size'
        )
    )
)


# ==========================================================
# 8. COMBINAR CLASES Y EXÁMENES
# ==========================================================

catalogo_bloques = resumen_clases.merge(
    resumen_examenes,
    on='BLOQUE',
    how='outer'
)

catalogo_bloques[
    'REGISTROS_CLASES'
] = (
    catalogo_bloques[
        'REGISTROS_CLASES'
    ]
    .fillna(0)
    .astype(int)
)

catalogo_bloques[
    'REGISTROS_EXAMENES'
] = (
    catalogo_bloques[
        'REGISTROS_EXAMENES'
    ]
    .fillna(0)
    .astype(int)
)

catalogo_bloques[
    'REGISTROS_TOTALES'
] = (
    catalogo_bloques[
        'REGISTROS_CLASES'
    ]
    +
    catalogo_bloques[
        'REGISTROS_EXAMENES'
    ]
)

catalogo_bloques[
    'REFERENCIAS'
] = catalogo_bloques.apply(
    combinar_referencias,
    axis=1
)

catalogo_bloques[
    'APARECE_EN_CLASES'
] = (
    catalogo_bloques[
        'REGISTROS_CLASES'
    ] > 0
).map({
    True: 'SI',
    False: 'NO'
})

catalogo_bloques[
    'APARECE_EN_EXAMENES'
] = (
    catalogo_bloques[
        'REGISTROS_EXAMENES'
    ] > 0
).map({
    True: 'SI',
    False: 'NO'
})

# El código del bloque será también su identificador.
catalogo_bloques.insert(
    0,
    'ID_BLOQUE',
    catalogo_bloques['BLOQUE']
)


# ==========================================================
# 9. COLUMNAS QUE SE COMPLETARÁN MANUALMENTE
# ==========================================================

# Solamente necesitamos una coordenada por bloque.
columnas_manuales = [
    'FACULTAD',
    'LATITUD',
    'LONGITUD'
]


# ==========================================================
# 10. CONSERVAR DATOS MANUALES EXISTENTES
# ==========================================================

if os.path.exists(ruta_catalogo_bloques):

    catalogo_anterior = pd.read_excel(
        ruta_catalogo_bloques,
        sheet_name='CATALOGO_BLOQUES'
    )

    if 'ID_BLOQUE' not in catalogo_anterior.columns:

        if 'BLOQUE' in catalogo_anterior.columns:

            catalogo_anterior[
                'ID_BLOQUE'
            ] = catalogo_anterior['BLOQUE']

        else:
            raise ValueError(
                'El catálogo anterior no tiene '
                'ID_BLOQUE ni BLOQUE.'
            )

    catalogo_anterior['ID_BLOQUE'] = (
        limpiar_texto(
            catalogo_anterior['ID_BLOQUE']
        )
    )

    columnas_anteriores_disponibles = [
        columna
        for columna in columnas_manuales
        if columna in catalogo_anterior.columns
    ]

    columnas_a_conservar = [
        'ID_BLOQUE'
    ] + columnas_anteriores_disponibles

    catalogo_anterior = (
        catalogo_anterior[
            columnas_a_conservar
        ]
        .drop_duplicates(
            subset=['ID_BLOQUE'],
            keep='last'
        )
    )

    catalogo_bloques = catalogo_bloques.merge(
        catalogo_anterior,
        on='ID_BLOQUE',
        how='left'
    )


# ==========================================================
# 11. CREAR COLUMNAS MANUALES FALTANTES
# ==========================================================

# Esto evita errores cuando el catálogo
# se ejecuta por primera vez.
for columna in columnas_manuales:

    if columna not in catalogo_bloques.columns:
        catalogo_bloques[columna] = pd.NA

catalogo_bloques['LATITUD'] = pd.to_numeric(
    catalogo_bloques['LATITUD'],
    errors='coerce'
)

catalogo_bloques['LONGITUD'] = pd.to_numeric(
    catalogo_bloques['LONGITUD'],
    errors='coerce'
)


# ==========================================================
# 12. ESTADO DE LAS COORDENADAS
# ==========================================================

catalogo_bloques[
    'ESTADO_COORDENADA'
] = 'PENDIENTE'

condicion_coordenada_completa = (
    catalogo_bloques['LATITUD'].notna()
    & catalogo_bloques['LONGITUD'].notna()
)

catalogo_bloques.loc[
    condicion_coordenada_completa,
    'ESTADO_COORDENADA'
] = 'COMPLETA'


# ==========================================================
# 13. CREAR RESUMEN
# ==========================================================

# El resumen debe crearse antes de eliminar
# las columnas auxiliares.
resumen = pd.DataFrame({
    'METRICA': [
        'BLOQUES UNICOS',
        'BLOQUES EN CLASES',
        'BLOQUES EN EXAMENES',
        'BLOQUES EN AMBAS FUENTES',
        'COORDENADAS COMPLETAS',
        'COORDENADAS PENDIENTES'
    ],

    'CANTIDAD': [
        len(catalogo_bloques),

        int(
            (
                catalogo_bloques[
                    'APARECE_EN_CLASES'
                ] == 'SI'
            ).sum()
        ),

        int(
            (
                catalogo_bloques[
                    'APARECE_EN_EXAMENES'
                ] == 'SI'
            ).sum()
        ),

        int(
            (
                (
                    catalogo_bloques[
                        'APARECE_EN_CLASES'
                    ] == 'SI'
                )
                &
                (
                    catalogo_bloques[
                        'APARECE_EN_EXAMENES'
                    ] == 'SI'
                )
            ).sum()
        ),

        int(
            (
                catalogo_bloques[
                    'ESTADO_COORDENADA'
                ] == 'COMPLETA'
            ).sum()
        ),

        int(
            (
                catalogo_bloques[
                    'ESTADO_COORDENADA'
                ] == 'PENDIENTE'
            ).sum()
        )
    ]
})


# ==========================================================
# 14. SELECCIONAR COLUMNAS FINALES
# ==========================================================

# Estas serán las únicas columnas visibles
# en la hoja CATALOGO_BLOQUES.
orden_columnas = [
    'ID_BLOQUE',
    'BLOQUE',
    'REFERENCIAS',
    'FACULTAD',
    'LATITUD',
    'LONGITUD'
]

catalogo_bloques = catalogo_bloques[
    orden_columnas
]


# ==========================================================
# 15. ORDENAR FILAS
# ==========================================================

catalogo_bloques[
    '_ORDEN_NUMERICO'
] = pd.to_numeric(
    catalogo_bloques['BLOQUE']
    .astype('string')
    .str.extract(
        r'(\d+)',
        expand=False
    ),
    errors='coerce'
)

catalogo_bloques = (
    catalogo_bloques
    .sort_values(
        [
            '_ORDEN_NUMERICO',
            'BLOQUE'
        ],
        na_position='last'
    )
    .drop(
        columns=['_ORDEN_NUMERICO']
    )
    .reset_index(drop=True)
)

# ==========================================================
# 16. EXPORTAR CATÁLOGO
# ==========================================================

with pd.ExcelWriter(
    ruta_catalogo_bloques,
    engine='openpyxl'
) as writer:

    catalogo_bloques.to_excel(
        writer,
        sheet_name='CATALOGO_BLOQUES',
        index=False
    )

    resumen.to_excel(
        writer,
        sheet_name='RESUMEN',
        index=False
    )


# ==========================================================
# 17. MOSTRAR RESULTADOS
# ==========================================================

print()
print('CATÁLOGO DE BLOQUES GENERADO')
print('----------------------------')

print(
    resumen.to_string(
        index=False
    )
)

print()
print(
    f'Archivo generado en:\n'
    f'{ruta_catalogo_bloques}'
)