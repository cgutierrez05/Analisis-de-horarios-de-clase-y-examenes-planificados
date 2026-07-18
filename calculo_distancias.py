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

ruta_catalogo_bloques = os.path.join(
    ruta_proyecto,
    'catalogos',
    'catalogo_bloques.xlsx'
)

ruta_catalogo_paradas = os.path.join(
    ruta_proyecto,
    'catalogos',
    'catalogo_paradas.xlsx'
)

ruta_resultado = os.path.join(
    ruta_proyecto,
    'catalogos',
    'relacion_bloque_parada.xlsx'
)

# Radio medio de la Tierra en metros.
RADIO_TIERRA_METROS = 6_371_000

# Cantidad de paradas cercanas que se conservarán
# por cada bloque y sentido de la ruta.
TOP_PARADAS = 3


# ==========================================================
# 2. FUNCIONES AUXILIARES
# ==========================================================

def normalizar_columnas(dataframe):
    """
    Normaliza los nombres de las columnas.

    Ejemplo:
    'Nombre Parada' -> 'NOMBRE_PARADA'
    """

    dataframe = dataframe.copy()

    dataframe.columns = (
        pd.Index(dataframe.columns)
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(
            r'\s+',
            '_',
            regex=True
        )
    )

    return dataframe


def limpiar_texto(serie):
    """
    Limpia espacios y convierte los textos a mayúsculas.
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


def convertir_coordenada(serie):
    """
    Convierte coordenadas a valores numéricos.

    También admite provisionalmente valores escritos
    con coma decimal, por ejemplo:
    -2,151899
    """

    return pd.to_numeric(
        serie
        .astype('string')
        .str.strip()
        .str.replace(
            ',',
            '.',
            regex=False
        ),
        errors='coerce'
    )


def leer_hoja_con_columnas(
    ruta_archivo,
    columnas_requeridas,
    hojas_preferidas=None
):
    """
    Busca dentro del Excel una hoja que contenga
    las columnas requeridas.

    Esto permite trabajar aunque la hoja se llame
    CATALOGO_BLOQUES, CATALOGO_PARADAS o Sheet1.
    """

    if not os.path.exists(ruta_archivo):
        raise FileNotFoundError(
            f'No se encontró el archivo:\n{ruta_archivo}'
        )

    archivo_excel = pd.ExcelFile(
        ruta_archivo
    )

    hojas = archivo_excel.sheet_names

    orden_busqueda = []

    if hojas_preferidas:

        for hoja in hojas_preferidas:

            if hoja in hojas:
                orden_busqueda.append(hoja)

    for hoja in hojas:

        if hoja not in orden_busqueda:
            orden_busqueda.append(hoja)

    for hoja in orden_busqueda:

        encabezado = pd.read_excel(
            ruta_archivo,
            sheet_name=hoja,
            nrows=0
        )

        encabezado = normalizar_columnas(
            encabezado
        )

        if columnas_requeridas.issubset(
            set(encabezado.columns)
        ):
            dataframe = pd.read_excel(
                ruta_archivo,
                sheet_name=hoja
            )

            dataframe = normalizar_columnas(
                dataframe
            )

            print(
                f'Hoja utilizada de '
                f'{os.path.basename(ruta_archivo)}: '
                f'{hoja}'
            )

            return dataframe

    raise ValueError(
        f'No se encontró una hoja válida en:\n'
        f'{ruta_archivo}\n'
        f'Columnas requeridas: '
        f'{sorted(columnas_requeridas)}'
    )


def calcular_distancia_haversine(
    latitud_origen,
    longitud_origen,
    latitud_destino,
    longitud_destino
):
    """
    Calcula la distancia en línea recta entre dos
    coordenadas geográficas utilizando Haversine.

    El resultado se expresa en metros.
    """

    latitud_origen_rad = np.radians(
        latitud_origen
    )

    longitud_origen_rad = np.radians(
        longitud_origen
    )

    latitud_destino_rad = np.radians(
        latitud_destino
    )

    longitud_destino_rad = np.radians(
        longitud_destino
    )

    diferencia_latitud = (
        latitud_destino_rad
        - latitud_origen_rad
    )

    diferencia_longitud = (
        longitud_destino_rad
        - longitud_origen_rad
    )

    valor_a = (
        np.sin(
            diferencia_latitud / 2
        ) ** 2
        +
        np.cos(latitud_origen_rad)
        * np.cos(latitud_destino_rad)
        * np.sin(
            diferencia_longitud / 2
        ) ** 2
    )

    # Evita errores numéricos mínimos.
    valor_a = np.clip(
        valor_a,
        0,
        1
    )

    valor_c = 2 * np.arctan2(
        np.sqrt(valor_a),
        np.sqrt(1 - valor_a)
    )

    return (
        RADIO_TIERRA_METROS
        * valor_c
    )


# ==========================================================
# 3. LEER CATÁLOGOS
# ==========================================================

columnas_bloques_requeridas = {
    'ID_BLOQUE',
    'BLOQUE',
    'REFERENCIAS',
    'LATITUD',
    'LONGITUD'
}

columnas_paradas_requeridas = {
    'ID_PARADA',
    'NOMBRE_PARADA',
    'LATITUD',
    'LONGITUD',
    'SENTIDO_RUTA'
}

bloques = leer_hoja_con_columnas(
    ruta_catalogo_bloques,
    columnas_bloques_requeridas,
    hojas_preferidas=[
        'CATALOGO_BLOQUES'
    ]
)

paradas = leer_hoja_con_columnas(
    ruta_catalogo_paradas,
    columnas_paradas_requeridas,
    hojas_preferidas=[
        'CATALOGO_PARADAS'
    ]
)

print()
print(
    f'Bloques leídos: {len(bloques)}'
)

print(
    f'Paradas leídas: {len(paradas)}'
)


# ==========================================================
# 4. LIMPIAR CATÁLOGO DE BLOQUES
# ==========================================================

bloques['ID_BLOQUE'] = limpiar_texto(
    bloques['ID_BLOQUE']
)

bloques['BLOQUE'] = limpiar_texto(
    bloques['BLOQUE']
)

bloques['REFERENCIAS'] = limpiar_texto(
    bloques['REFERENCIAS']
)

bloques['LATITUD'] = convertir_coordenada(
    bloques['LATITUD']
)

bloques['LONGITUD'] = convertir_coordenada(
    bloques['LONGITUD']
)


# ==========================================================
# 5. LIMPIAR CATÁLOGO DE PARADAS
# ==========================================================

paradas['ID_PARADA'] = limpiar_texto(
    paradas['ID_PARADA']
)

paradas['NOMBRE_PARADA'] = limpiar_texto(
    paradas['NOMBRE_PARADA']
)

paradas['SENTIDO_RUTA'] = limpiar_texto(
    paradas['SENTIDO_RUTA']
)

paradas['LATITUD'] = convertir_coordenada(
    paradas['LATITUD']
)

paradas['LONGITUD'] = convertir_coordenada(
    paradas['LONGITUD']
)


# ==========================================================
# 6. VALIDAR IDENTIFICADORES
# ==========================================================

if bloques['ID_BLOQUE'].isna().any():
    raise ValueError(
        'Existen bloques sin ID_BLOQUE.'
    )

if paradas['ID_PARADA'].isna().any():
    raise ValueError(
        'Existen paradas sin ID_PARADA.'
    )

duplicados_bloques = bloques[
    bloques['ID_BLOQUE'].duplicated(
        keep=False
    )
].copy()

duplicados_paradas = paradas[
    paradas['ID_PARADA'].duplicated(
        keep=False
    )
].copy()

if not duplicados_bloques.empty:
    raise ValueError(
        'Existen ID_BLOQUE duplicados:\n'
        + duplicados_bloques[
            [
                'ID_BLOQUE',
                'BLOQUE'
            ]
        ].to_string(index=False)
    )

if not duplicados_paradas.empty:
    raise ValueError(
        'Existen ID_PARADA duplicados:\n'
        + duplicados_paradas[
            [
                'ID_PARADA',
                'NOMBRE_PARADA'
            ]
        ].to_string(index=False)
    )


# ==========================================================
# 7. VALIDAR COORDENADAS
# ==========================================================

bloques['FLAG_COORDENADA_INCOMPLETA'] = (
    bloques['LATITUD'].isna()
    | bloques['LONGITUD'].isna()
).astype(int)

bloques['FLAG_COORDENADA_FUERA_RANGO'] = (
    (
        bloques['LATITUD'].notna()
        & ~bloques['LATITUD'].between(
            -90,
            90
        )
    )
    |
    (
        bloques['LONGITUD'].notna()
        & ~bloques['LONGITUD'].between(
            -180,
            180
        )
    )
).astype(int)

paradas['FLAG_COORDENADA_INCOMPLETA'] = (
    paradas['LATITUD'].isna()
    | paradas['LONGITUD'].isna()
).astype(int)

paradas['FLAG_COORDENADA_FUERA_RANGO'] = (
    (
        paradas['LATITUD'].notna()
        & ~paradas['LATITUD'].between(
            -90,
            90
        )
    )
    |
    (
        paradas['LONGITUD'].notna()
        & ~paradas['LONGITUD'].between(
            -180,
            180
        )
    )
).astype(int)


# ==========================================================
# 8. VALIDAR SENTIDO DE LA RUTA
# ==========================================================

sentidos_validos = [
    'ENTRADA',
    'SALIDA'
]

paradas['FLAG_SENTIDO_INVALIDO'] = (
    paradas['SENTIDO_RUTA'].isna()
    |
    ~paradas['SENTIDO_RUTA'].isin(
        sentidos_validos
    )
).astype(int)


# ==========================================================
# 9. SEPARAR REGISTROS VÁLIDOS Y PENDIENTES
# ==========================================================

bloques_validos = bloques[
    (bloques['FLAG_COORDENADA_INCOMPLETA'] == 0)
    & (
        bloques[
            'FLAG_COORDENADA_FUERA_RANGO'
        ] == 0
    )
].copy()

bloques_pendientes = bloques[
    ~bloques['ID_BLOQUE'].isin(
        bloques_validos['ID_BLOQUE']
    )
].copy()

paradas_validas = paradas[
    (paradas['FLAG_COORDENADA_INCOMPLETA'] == 0)
    & (
        paradas[
            'FLAG_COORDENADA_FUERA_RANGO'
        ] == 0
    )
    & (
        paradas[
            'FLAG_SENTIDO_INVALIDO'
        ] == 0
    )
].copy()

paradas_pendientes = paradas[
    ~paradas['ID_PARADA'].isin(
        paradas_validas['ID_PARADA']
    )
].copy()


# ==========================================================
# 10. COMPROBAR QUE EXISTAN DATOS PARA CALCULAR
# ==========================================================

if bloques_validos.empty:
    print()
    print(
        'No existen bloques con coordenadas completas.'
    )

if paradas_validas.empty:
    print()
    print(
        'No existen paradas válidas para calcular.'
    )


# ==========================================================
# 11. GENERAR TODAS LAS COMBINACIONES BLOQUE-PARADA
# ==========================================================

if (
    not bloques_validos.empty
    and not paradas_validas.empty
):

    bloques_para_cruce = bloques_validos[
        [
            'ID_BLOQUE',
            'BLOQUE',
            'REFERENCIAS',
            'LATITUD',
            'LONGITUD'
        ]
    ].rename(
        columns={
            'LATITUD': 'LATITUD_BLOQUE',
            'LONGITUD': 'LONGITUD_BLOQUE'
        }
    )

    paradas_para_cruce = paradas_validas[
        [
            'ID_PARADA',
            'NOMBRE_PARADA',
            'SENTIDO_RUTA',
            'LATITUD',
            'LONGITUD'
        ]
    ].rename(
        columns={
            'LATITUD': 'LATITUD_PARADA',
            'LONGITUD': 'LONGITUD_PARADA'
        }
    )

    distancias = bloques_para_cruce.merge(
        paradas_para_cruce,
        how='cross'
    )


    # ======================================================
    # 12. CALCULAR DISTANCIA HAVERSINE
    # ======================================================

    distancias['DISTANCIA_METROS'] = (
        calcular_distancia_haversine(
            distancias['LATITUD_BLOQUE'],
            distancias['LONGITUD_BLOQUE'],
            distancias['LATITUD_PARADA'],
            distancias['LONGITUD_PARADA']
        )
        .round(2)
    )


    # ======================================================
    # 13. ORDENAR POR CERCANÍA
    # ======================================================

    distancias = distancias.sort_values(
        [
            'ID_BLOQUE',
            'SENTIDO_RUTA',
            'DISTANCIA_METROS',
            'ID_PARADA'
        ]
    ).reset_index(drop=True)

    # Orden dentro de cada sentido:
    # entrada y salida se analizan por separado.
    distancias[
        'ORDEN_CERCANIA_SENTIDO'
    ] = (
        distancias
        .groupby(
            [
                'ID_BLOQUE',
                'SENTIDO_RUTA'
            ]
        )
        .cumcount()
        + 1
    )

    # Orden general sin distinguir sentido.
    distancias_orden_general = (
        distancias
        .sort_values(
            [
                'ID_BLOQUE',
                'DISTANCIA_METROS',
                'ID_PARADA'
            ]
        )
        .copy()
    )

    distancias_orden_general[
        'ORDEN_CERCANIA_GENERAL'
    ] = (
        distancias_orden_general
        .groupby('ID_BLOQUE')
        .cumcount()
        + 1
    )

    distancias = distancias.merge(
        distancias_orden_general[
            [
                'ID_BLOQUE',
                'ID_PARADA',
                'ORDEN_CERCANIA_GENERAL'
            ]
        ],
        on=[
            'ID_BLOQUE',
            'ID_PARADA'
        ],
        how='left',
        validate='one_to_one'
    )


    # ======================================================
    # 14. CONSERVAR LAS TRES MÁS CERCANAS
    # ======================================================

    top_paradas = distancias[
        distancias[
            'ORDEN_CERCANIA_SENTIDO'
        ] <= TOP_PARADAS
    ].copy()

    parada_principal = distancias[
        distancias[
            'ORDEN_CERCANIA_SENTIDO'
        ] == 1
    ].copy()


    # ======================================================
    # 15. CREAR ASIGNACIÓN PRINCIPAL POR BLOQUE
    # ======================================================

    asignacion_entrada = parada_principal[
        parada_principal[
            'SENTIDO_RUTA'
        ] == 'ENTRADA'
    ][
        [
            'ID_BLOQUE',
            'ID_PARADA',
            'NOMBRE_PARADA',
            'DISTANCIA_METROS'
        ]
    ].rename(
        columns={
            'ID_PARADA':
                'ID_PARADA_ENTRADA',
            'NOMBRE_PARADA':
                'NOMBRE_PARADA_ENTRADA',
            'DISTANCIA_METROS':
                'DISTANCIA_ENTRADA_METROS'
        }
    )

    asignacion_salida = parada_principal[
        parada_principal[
            'SENTIDO_RUTA'
        ] == 'SALIDA'
    ][
        [
            'ID_BLOQUE',
            'ID_PARADA',
            'NOMBRE_PARADA',
            'DISTANCIA_METROS'
        ]
    ].rename(
        columns={
            'ID_PARADA':
                'ID_PARADA_SALIDA',
            'NOMBRE_PARADA':
                'NOMBRE_PARADA_SALIDA',
            'DISTANCIA_METROS':
                'DISTANCIA_SALIDA_METROS'
        }
    )

    asignacion_principal = bloques_validos[
        [
            'ID_BLOQUE',
            'BLOQUE',
            'REFERENCIAS',
            'LATITUD',
            'LONGITUD'
        ]
    ].copy()

    asignacion_principal = (
        asignacion_principal
        .merge(
            asignacion_entrada,
            on='ID_BLOQUE',
            how='left',
            validate='one_to_one'
        )
        .merge(
            asignacion_salida,
            on='ID_BLOQUE',
            how='left',
            validate='one_to_one'
        )
    )

else:

    columnas_distancias = [
        'ID_BLOQUE',
        'BLOQUE',
        'REFERENCIAS',
        'LATITUD_BLOQUE',
        'LONGITUD_BLOQUE',
        'ID_PARADA',
        'NOMBRE_PARADA',
        'SENTIDO_RUTA',
        'LATITUD_PARADA',
        'LONGITUD_PARADA',
        'DISTANCIA_METROS',
        'ORDEN_CERCANIA_SENTIDO',
        'ORDEN_CERCANIA_GENERAL'
    ]

    distancias = pd.DataFrame(
        columns=columnas_distancias
    )

    top_paradas = distancias.copy()

    parada_principal = distancias.copy()

    asignacion_principal = pd.DataFrame(
        columns=[
            'ID_BLOQUE',
            'BLOQUE',
            'REFERENCIAS',
            'LATITUD',
            'LONGITUD',
            'ID_PARADA_ENTRADA',
            'NOMBRE_PARADA_ENTRADA',
            'DISTANCIA_ENTRADA_METROS',
            'ID_PARADA_SALIDA',
            'NOMBRE_PARADA_SALIDA',
            'DISTANCIA_SALIDA_METROS'
        ]
    )


# ==========================================================
# 16. CREAR RESUMEN
# ==========================================================

resumen = pd.DataFrame({
    'METRICA': [
        'BLOQUES TOTALES',
        'BLOQUES CON COORDENADAS VALIDAS',
        'BLOQUES PENDIENTES',
        'PARADAS TOTALES',
        'PARADAS VALIDAS',
        'PARADAS PENDIENTES',
        'PARADAS DE ENTRADA VALIDAS',
        'PARADAS DE SALIDA VALIDAS',
        'COMBINACIONES CALCULADAS',
        'ASIGNACIONES PRINCIPALES',
        'TOP DE PARADAS CONSERVADAS'
    ],

    'CANTIDAD': [
        len(bloques),
        len(bloques_validos),
        len(bloques_pendientes),
        len(paradas),
        len(paradas_validas),
        len(paradas_pendientes),

        int(
            (
                paradas_validas[
                    'SENTIDO_RUTA'
                ] == 'ENTRADA'
            ).sum()
        ),

        int(
            (
                paradas_validas[
                    'SENTIDO_RUTA'
                ] == 'SALIDA'
            ).sum()
        ),

        len(distancias),
        len(parada_principal),
        len(top_paradas)
    ]
})


# ==========================================================
# 17. EXPORTAR RESULTADOS
# ==========================================================

os.makedirs(
    os.path.dirname(ruta_resultado),
    exist_ok=True
)

with pd.ExcelWriter(
    ruta_resultado,
    engine='openpyxl'
) as writer:

    asignacion_principal.to_excel(
        writer,
        sheet_name='ASIGNACION_PRINCIPAL',
        index=False
    )

    top_paradas.to_excel(
        writer,
        sheet_name='TOP_3_POR_SENTIDO',
        index=False
    )

    distancias.to_excel(
        writer,
        sheet_name='TODAS_DISTANCIAS',
        index=False
    )

    bloques_pendientes.to_excel(
        writer,
        sheet_name='BLOQUES_PENDIENTES',
        index=False
    )

    paradas_pendientes.to_excel(
        writer,
        sheet_name='PARADAS_PENDIENTES',
        index=False
    )

    resumen.to_excel(
        writer,
        sheet_name='RESUMEN',
        index=False
    )


# ==========================================================
# 18. MOSTRAR RESULTADOS
# ==========================================================

print()
print('CÁLCULO DE DISTANCIAS FINALIZADO')
print('--------------------------------')

print(
    resumen.to_string(
        index=False
    )
)

print()
print(
    f'Archivo generado en:\n'
    f'{ruta_resultado}'
)