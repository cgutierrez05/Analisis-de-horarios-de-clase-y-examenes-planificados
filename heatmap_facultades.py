import os
import unicodedata
from datetime import datetime, time

import matplotlib.pyplot as plt
import numpy as np
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

ruta_catalogo_bloques = os.path.join(
    ruta_proyecto,
    'catalogos',
    'catalogo_bloques.xlsx'
)

ruta_resultados = os.path.join(
    ruta_proyecto,
    'resultados',
    'heatmap_facultades'
)

ruta_excel_resultado = os.path.join(
    ruta_resultados,
    'concentracion_estudiantes_facultad.xlsx'
)

os.makedirs(
    ruta_resultados,
    exist_ok=True
)


# ==========================================================
# 2. PARÁMETROS DEL ANÁLISIS
# ==========================================================

franjas_horarias = [
    {
        'FRANJA': '09:00-11:00',
        'INICIO': 9 * 60,
        'FIN': 11 * 60
    },
    {
        'FRANJA': '11:00-13:00',
        'INICIO': 11 * 60,
        'FIN': 13 * 60
    },
    {
        'FRANJA': '13:00-15:00',
        'INICIO': 13 * 60,
        'FIN': 15 * 60
    },
    {
        'FRANJA': '15:00-17:00',
        'INICIO': 15 * 60,
        'FIN': 17 * 60
    }
]

orden_franjas = [
    '09:00-11:00',
    '11:00-13:00',
    '13:00-15:00',
    '15:00-17:00'
]

orden_dias = [
    'LUNES',
    'MARTES',
    'MIERCOLES',
    'JUEVES',
    'VIERNES'
]

duracion_franja_minutos = 120


# ==========================================================
# 3. FUNCIONES AUXILIARES
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


def quitar_tildes(valor):
    """
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
    Convierte diferentes formatos de hora a minutos
    transcurridos desde las 00:00.
    """

    if pd.isna(valor):
        return pd.NA

    if isinstance(valor, (pd.Timestamp, datetime)):
        return (
            valor.hour * 60
            + valor.minute
        )

    if isinstance(valor, time):
        return (
            valor.hour * 60
            + valor.minute
        )

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

        # Fracción horaria de Excel.
        if 0 <= numero < 1:
            return round(
                numero * 1440
            ) % 1440

        return round(
            (numero % 1) * 1440
        ) % 1440

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


def calcular_superposicion(
    inicio_clase,
    fin_clase,
    inicio_franja,
    fin_franja
):
    """
    Calcula cuántos minutos de una clase se encuentran
    dentro de una franja horaria.
    """

    if (
        pd.isna(inicio_clase)
        or pd.isna(fin_clase)
    ):
        return 0

    inicio_superposicion = max(
        int(inicio_clase),
        int(inicio_franja)
    )

    fin_superposicion = min(
        int(fin_clase),
        int(fin_franja)
    )

    return max(
        0,
        fin_superposicion
        - inicio_superposicion
    )


def guardar_heatmap(
    matriz,
    titulo,
    ruta_imagen,
    etiqueta_barra
):
    """
    Guarda una matriz como mapa de calor.
    """

    if matriz.empty:
        print(
            f'No se generó "{titulo}" porque '
            'la matriz está vacía.'
        )
        return

    ancho = max(
        8,
        len(matriz.columns) * 1.8
    )

    alto = max(
        4,
        len(matriz.index) * 0.65
    )

    figura, eje = plt.subplots(
        figsize=(ancho, alto)
    )

    imagen = eje.imshow(
        matriz.values,
        aspect='auto'
    )

    eje.set_xticks(
        range(len(matriz.columns))
    )

    eje.set_xticklabels(
        matriz.columns,
        rotation=35,
        ha='right'
    )

    eje.set_yticks(
        range(len(matriz.index))
    )

    eje.set_yticklabels(
        matriz.index
    )

    eje.set_xlabel(
        'Franja horaria'
    )

    eje.set_ylabel(
        'Facultad'
    )

    eje.set_title(
        titulo
    )

    # Mostrar el valor dentro de cada celda.
    for fila in range(len(matriz.index)):

        for columna in range(
            len(matriz.columns)
        ):

            valor = matriz.iloc[
                fila,
                columna
            ]

            eje.text(
                columna,
                fila,
                f'{valor:.0f}',
                ha='center',
                va='center'
            )

    barra = figura.colorbar(
        imagen,
        ax=eje
    )

    barra.set_label(
        etiqueta_barra
    )

    figura.tight_layout()

    figura.savefig(
        ruta_imagen,
        dpi=300,
        bbox_inches='tight'
    )

    plt.close(figura)


# ==========================================================
# 4. COMPROBAR ARCHIVOS
# ==========================================================

if not os.path.exists(ruta_clases):
    raise FileNotFoundError(
        'No se encontró el archivo de clases:\n'
        f'{ruta_clases}'
    )

if not os.path.exists(ruta_catalogo_bloques):
    raise FileNotFoundError(
        'No se encontró el catálogo de bloques:\n'
        f'{ruta_catalogo_bloques}'
    )


# ==========================================================
# 5. LEER DATOS
# ==========================================================

df_clases = pd.read_excel(
    ruta_clases,
    sheet_name='DATOS_LIMPIOS'
)

df_bloques = pd.read_excel(
    ruta_catalogo_bloques,
    sheet_name='CATALOGO_BLOQUES'
)

print(
    f'Registros limpios de clases: '
    f'{len(df_clases)}'
)

print(
    f'Bloques en el catálogo: '
    f'{len(df_bloques)}'
)


# ==========================================================
# 6. VALIDAR COLUMNAS
# ==========================================================

columnas_clases_necesarias = {
    'BLOQUE',
    'DIA',
    'HORAINICIO',
    'HORAFIN',
    'NUMREGISTRADOS_AJUSTADO'
}

columnas_bloques_necesarias = {
    'BLOQUE',
    'FACULTAD'
}

faltantes_clases = (
    columnas_clases_necesarias
    - set(df_clases.columns)
)

faltantes_bloques = (
    columnas_bloques_necesarias
    - set(df_bloques.columns)
)

if faltantes_clases:
    raise ValueError(
        'Faltan columnas en DATOS_LIMPIOS: '
        + ', '.join(
            sorted(faltantes_clases)
        )
    )

if faltantes_bloques:
    raise ValueError(
        'Faltan columnas en CATALOGO_BLOQUES: '
        + ', '.join(
            sorted(faltantes_bloques)
        )
    )


# ==========================================================
# 7. LIMPIAR DATOS
# ==========================================================

df_clases['BLOQUE'] = limpiar_texto(
    df_clases['BLOQUE']
)

df_bloques['BLOQUE'] = limpiar_texto(
    df_bloques['BLOQUE']
)

df_bloques['FACULTAD'] = limpiar_texto(
    df_bloques['FACULTAD']
)

df_clases['DIA_LIMPIO'] = (
    df_clases['DIA']
    .map(quitar_tildes)
    .astype('string')
    .str.upper()
)

df_clases['MINUTO_INICIO'] = (
    df_clases['HORAINICIO']
    .map(hora_a_minutos)
    .astype('Int64')
)

df_clases['MINUTO_FIN'] = (
    df_clases['HORAFIN']
    .map(hora_a_minutos)
    .astype('Int64')
)

df_clases[
    'NUMREGISTRADOS_AJUSTADO'
] = pd.to_numeric(
    df_clases[
        'NUMREGISTRADOS_AJUSTADO'
    ],
    errors='coerce'
)


# ==========================================================
# 8. VALIDAR CATÁLOGO DE BLOQUES
# ==========================================================

duplicados_bloques = df_bloques[
    df_bloques['BLOQUE'].duplicated(
        keep=False
    )
].copy()

if not duplicados_bloques.empty:

    raise ValueError(
        'Existen bloques duplicados en el catálogo:\n'
        + duplicados_bloques[
            [
                'BLOQUE',
                'FACULTAD'
            ]
        ].to_string(index=False)
    )


# ==========================================================
# 9. UNIR CLASES CON FACULTADES
# ==========================================================

df_clases = df_clases.merge(
    df_bloques[
        [
            'BLOQUE',
            'FACULTAD'
        ]
    ],
    on='BLOQUE',
    how='left',
    validate='many_to_one'
)

df_bloques_sin_facultad = (
    df_clases[
        df_clases['FACULTAD'].isna()
    ][
        [
            'BLOQUE'
        ]
    ]
    .drop_duplicates()
    .sort_values('BLOQUE')
    .reset_index(drop=True)
)

if not df_bloques_sin_facultad.empty:

    print()
    print(
        'ADVERTENCIA: existen bloques sin facultad:'
    )

    print(
        df_bloques_sin_facultad.to_string(
            index=False
        )
    )

# Solo se incluyen en el heatmap los bloques
# con una facultad identificada.
df_clases_analisis = df_clases[
    df_clases['FACULTAD'].notna()
    & df_clases[
        'NUMREGISTRADOS_AJUSTADO'
    ].gt(0)
    & df_clases['MINUTO_INICIO'].notna()
    & df_clases['MINUTO_FIN'].notna()
    & (
        df_clases['MINUTO_FIN']
        > df_clases['MINUTO_INICIO']
    )
].copy()

# Si todavía existe la bandera física,
# se conservan únicamente eventos físicos.
if 'ES_EVENTO_FISICO' in df_clases_analisis.columns:

    df_clases_analisis = (
        df_clases_analisis[
            df_clases_analisis[
                'ES_EVENTO_FISICO'
            ] == 1
        ]
        .copy()
    )


# ==========================================================
# 10. CALCULAR APORTE A CADA FRANJA
# ==========================================================

registros_intervalos = []

for _, fila in df_clases_analisis.iterrows():

    for franja in franjas_horarias:

        minutos_superpuestos = (
            calcular_superposicion(
                fila['MINUTO_INICIO'],
                fila['MINUTO_FIN'],
                franja['INICIO'],
                franja['FIN']
            )
        )

        if minutos_superpuestos <= 0:
            continue

        estudiantes_promedio = (
            fila[
                'NUMREGISTRADOS_AJUSTADO'
            ]
            * minutos_superpuestos
            / duracion_franja_minutos
        )

        registros_intervalos.append({
            'ID_FILA_ORIGINAL': fila.get(
                'ID_FILA_ORIGINAL',
                pd.NA
            ),
            'DIA': fila['DIA_LIMPIO'],
            'FACULTAD': fila['FACULTAD'],
            'BLOQUE': fila['BLOQUE'],
            'FRANJA': franja['FRANJA'],
            'MINUTOS_SUPERPUESTOS':
                minutos_superpuestos,
            'NUMREGISTRADOS_AJUSTADO':
                fila[
                    'NUMREGISTRADOS_AJUSTADO'
                ],
            'ESTUDIANTES_PROMEDIO_FRANJA':
                estudiantes_promedio
        })


df_intervalos = pd.DataFrame(
    registros_intervalos
)

if df_intervalos.empty:
    raise ValueError(
        'No se encontraron clases dentro de las '
        'franjas de 09:00 a 17:00.'
    )


# ==========================================================
# 11. RESUMEN POR DÍA, FACULTAD Y FRANJA
# ==========================================================

resumen_por_dia = (
    df_intervalos
    .groupby(
        [
            'DIA',
            'FACULTAD',
            'FRANJA'
        ],
        as_index=False
    )
    .agg(
        ESTUDIANTES_PROMEDIO=(
            'ESTUDIANTES_PROMEDIO_FRANJA',
            'sum'
        ),
        MINUTOS_CLASE_ACUMULADOS=(
            'MINUTOS_SUPERPUESTOS',
            'sum'
        ),
        EVENTOS_ACADEMICOS=(
            'ID_FILA_ORIGINAL',
            'count'
        )
    )
)

resumen_por_dia[
    'ESTUDIANTES_PROMEDIO'
] = resumen_por_dia[
    'ESTUDIANTES_PROMEDIO'
].round(2)


# ==========================================================
# 12. CREAR CUADRÍCULA COMPLETA
# ==========================================================

facultades = sorted(
    df_intervalos[
        'FACULTAD'
    ].dropna().unique()
)

cuadricula_completa = pd.MultiIndex.from_product(
    [
        orden_dias,
        facultades,
        orden_franjas
    ],
    names=[
        'DIA',
        'FACULTAD',
        'FRANJA'
    ]
).to_frame(index=False)

resumen_completo = cuadricula_completa.merge(
    resumen_por_dia,
    on=[
        'DIA',
        'FACULTAD',
        'FRANJA'
    ],
    how='left'
)

columnas_numericas = [
    'ESTUDIANTES_PROMEDIO',
    'MINUTOS_CLASE_ACUMULADOS',
    'EVENTOS_ACADEMICOS'
]

resumen_completo[
    columnas_numericas
] = resumen_completo[
    columnas_numericas
].fillna(0)


# ==========================================================
# 13. PROMEDIO DE LUNES A VIERNES
# ==========================================================

promedio_semanal = (
    resumen_completo
    .groupby(
        [
            'FACULTAD',
            'FRANJA'
        ],
        as_index=False
    )
    .agg(
        ESTUDIANTES_PROMEDIO_DIA=(
            'ESTUDIANTES_PROMEDIO',
            'mean'
        )
    )
)

promedio_semanal[
    'ESTUDIANTES_PROMEDIO_DIA'
] = promedio_semanal[
    'ESTUDIANTES_PROMEDIO_DIA'
].round(2)


# ==========================================================
# 14. TOTAL ACUMULADO DE LA SEMANA
# ==========================================================

total_semanal = (
    resumen_completo
    .groupby(
        [
            'FACULTAD',
            'FRANJA'
        ],
        as_index=False
    )
    .agg(
        ESTUDIANTES_ACUMULADOS_SEMANA=(
            'ESTUDIANTES_PROMEDIO',
            'sum'
        )
    )
)

total_semanal[
    'ESTUDIANTES_ACUMULADOS_SEMANA'
] = total_semanal[
    'ESTUDIANTES_ACUMULADOS_SEMANA'
].round(2)


# ==========================================================
# 15. MATRIZ PROMEDIO SEMANAL
# ==========================================================

matriz_promedio = (
    promedio_semanal
    .pivot(
        index='FACULTAD',
        columns='FRANJA',
        values='ESTUDIANTES_PROMEDIO_DIA'
    )
    .reindex(
        columns=orden_franjas
    )
    .fillna(0)
)

guardar_heatmap(
    matriz=matriz_promedio,
    titulo=(
        'Concentración promedio de estudiantes '
        'por facultad'
    ),
    ruta_imagen=os.path.join(
        ruta_resultados,
        'heatmap_promedio_semanal.png'
    ),
    etiqueta_barra=(
        'Estudiantes promedio estimados por día'
    )
)


# ==========================================================
# 16. MATRIZ TOTAL SEMANAL
# ==========================================================

matriz_total = (
    total_semanal
    .pivot(
        index='FACULTAD',
        columns='FRANJA',
        values='ESTUDIANTES_ACUMULADOS_SEMANA'
    )
    .reindex(
        columns=orden_franjas
    )
    .fillna(0)
)

guardar_heatmap(
    matriz=matriz_total,
    titulo=(
        'Concentración acumulada semanal '
        'por facultad'
    ),
    ruta_imagen=os.path.join(
        ruta_resultados,
        'heatmap_total_semanal.png'
    ),
    etiqueta_barra=(
        'Estudiantes estimados acumulados'
    )
)


# ==========================================================
# 17. HEATMAP INDIVIDUAL POR DÍA
# ==========================================================

for dia in orden_dias:

    datos_dia = resumen_completo[
        resumen_completo['DIA'] == dia
    ]

    matriz_dia = (
        datos_dia
        .pivot(
            index='FACULTAD',
            columns='FRANJA',
            values='ESTUDIANTES_PROMEDIO'
        )
        .reindex(
            index=facultades,
            columns=orden_franjas
        )
        .fillna(0)
    )

    guardar_heatmap(
        matriz=matriz_dia,
        titulo=(
            f'Concentración de estudiantes - {dia}'
        ),
        ruta_imagen=os.path.join(
            ruta_resultados,
            f'heatmap_{dia.lower()}.png'
        ),
        etiqueta_barra=(
            'Estudiantes promedio estimados'
        )
    )


# ==========================================================
# 18. EXPORTAR RESULTADOS A EXCEL
# ==========================================================

with pd.ExcelWriter(
    ruta_excel_resultado,
    engine='openpyxl'
) as writer:

    df_intervalos.to_excel(
        writer,
        sheet_name='DETALLE_INTERVALOS',
        index=False
    )

    resumen_completo.to_excel(
        writer,
        sheet_name='RESUMEN_POR_DIA',
        index=False
    )

    promedio_semanal.to_excel(
        writer,
        sheet_name='PROMEDIO_SEMANAL',
        index=False
    )

    total_semanal.to_excel(
        writer,
        sheet_name='TOTAL_SEMANAL',
        index=False
    )

    matriz_promedio.to_excel(
        writer,
        sheet_name='MATRIZ_PROMEDIO'
    )

    matriz_total.to_excel(
        writer,
        sheet_name='MATRIZ_TOTAL'
    )

    df_bloques_sin_facultad.to_excel(
        writer,
        sheet_name='BLOQUES_SIN_FACULTAD',
        index=False
    )


# ==========================================================
# 19. MOSTRAR RESULTADOS
# ==========================================================

print()
print('HEATMAPS GENERADOS')
print('------------------')

print(
    f'Registros analizados: '
    f'{len(df_clases_analisis)}'
)

print(
    f'Facultades analizadas: '
    f'{len(facultades)}'
)

print(
    f'Bloques sin facultad: '
    f'{len(df_bloques_sin_facultad)}'
)

print()
print(
    f'Resultados guardados en:\n'
    f'{ruta_resultados}'
)

print()
print(
    'Archivos principales:'
)

print(
    '- heatmap_promedio_semanal.png'
)

print(
    '- heatmap_total_semanal.png'
)

print(
    '- heatmap_lunes.png hasta heatmap_viernes.png'
)

print(
    '- concentracion_estudiantes_facultad.xlsx'
)