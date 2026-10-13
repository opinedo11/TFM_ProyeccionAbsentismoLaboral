#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
CONSTRUCCIÓN DE LA TABLA DE ENTRENAMIENTO
PERSONA-SEMANA + HISTÓRICO DE IT

VERSIÓN 4

Cambios principales respecto a versiones anteriores:
    - NO utiliza la variable TipoJornada.
    - NO utiliza ponderaciones 5/7.
    - NO genera valores decimales.
    - El cálculo se realiza a nivel de día calendario.
    - Para cada empleado y semana se seleccionan como máximo
      5 días computables de IT.
    - Total = ShortTerm + LongTerm en todas las filas.
    - Las ventanas históricas se calculan sobre los días
      semanales consolidados, respetando exactamente sus límites.
    - Incluye controles específicos para el empleado 23109.

Regla de días:
    1. Se genera un registro por empleado, episodio y día.
    2. Cada día recibe la clasificación ShortTerm o LongTerm
       según la duración total del episodio.
    3. Para cada empleado-semana se toman como máximo 5 días,
       ordenados cronológicamente.
    4. Los días seleccionados conservan su clasificación.
    5. Por tanto:
           Total = ShortTerm + LongTerm
       siempre.

Ejemplo 23109:
    IT: 19/05/2024 - 25/05/2024

    Semana 20:
        19/05 -> 1 día LongTerm

    Semana 21:
        20/05 - 25/05 -> 6 días calendario
        -> se conservan 5

    Resultado:
        Semana 20 = 1
        Semana 21 = 5
        Total = 6

Bajas abiertas:
    - ValidTo = 99991231 se interpreta como abierta.
    - Para construir el histórico se considera la parte
      transcurrida hasta la fecha máxima de corte.
    - La duración utilizada para clasificar un episodio abierto
      es su duración observada hasta dicho corte.
    - El empleado permanece no elegible mientras la baja está abierta.
"""

from pathlib import Path
import json

import openpyxl
import duckdb


# ============================================================
# CONFIGURACIÓN
# ============================================================

CARPETA_BASE = Path(__file__).resolve().parent

ARCHIVO_PERSONA = CARPETA_BASE / "Persona_Semana.csv"
ARCHIVO_IT = CARPETA_BASE / "AbsentismosIT.xlsx"

ARCHIVO_SALIDA = (
    CARPETA_BASE / "Tabla_Entrenamiento_IT.csv"
)

ARCHIVO_RESUMEN = (
    CARPETA_BASE / "ResumenTablaEntrenamiento.json"
)

FECHA_CENTINELA = "99991231"


def ruta_sql(ruta):
    return (
        str(ruta)
        .replace("\\", "/")
        .replace("'", "''")
    )


# ============================================================
# CONEXIÓN
# ============================================================

con = duckdb.connect()


# ============================================================
# 1. PERSONA-SEMANA
# ============================================================

print("1/9 - Cargando Persona_Semana.csv...")

if not ARCHIVO_PERSONA.exists():
    raise FileNotFoundError(
        f"No se encontró: {ARCHIVO_PERSONA}"
    )

con.execute(f"""
    CREATE OR REPLACE TABLE persona AS
    SELECT *
    FROM read_csv(
        '{ruta_sql(ARCHIVO_PERSONA)}',
        delim=';',
        header=true,
        auto_detect=true,
        encoding='UTF-8',
        null_padding=true,
        ignore_errors=false
    )
""")

columnas_persona = {
    fila[0]
    for fila in con.execute(
        "DESCRIBE persona"
    ).fetchall()
}

requeridas = {
    "CodEmpleado",
    "FechaInicio",
    "FechaFin",
    "AnioISO",
    "SemanaISO",
    "ID_Semana",
    "Area",
    "Servicio"
}

faltantes = requeridas - columnas_persona

if faltantes:
    raise ValueError(
        "Faltan columnas en Persona_Semana.csv: "
        f"{sorted(faltantes)}"
    )

con.execute("""
    CREATE OR REPLACE TABLE persona_base AS
    SELECT
        *,
        trim(
            cast("CodEmpleado" AS VARCHAR)
        ) AS empleado_key,
        try_cast(
            "FechaInicio" AS DATE
        ) AS inicio_semana,
        try_cast(
            "FechaFin" AS DATE
        ) AS fin_semana
    FROM persona
    WHERE "CodEmpleado" IS NOT NULL
""")

total_persona = con.execute(
    "SELECT count(*) FROM persona_base"
).fetchone()[0]

print(f"   Filas: {total_persona:,}")


# ============================================================
# 2. IT
# ============================================================

print("2/9 - Cargando AbsentismosIT.xlsx...")

if not ARCHIVO_IT.exists():
    raise FileNotFoundError(
        f"No se encontró: {ARCHIVO_IT}"
    )

libro = openpyxl.load_workbook(
    ARCHIVO_IT,
    read_only=True,
    data_only=True
)

hoja = libro.active
filas = hoja.iter_rows(values_only=True)

encabezados = next(filas, None)

if encabezados is None:
    raise ValueError(
        "AbsentismosIT.xlsx no contiene datos."
    )

encabezados = [
    str(x).strip() if x is not None else ""
    for x in encabezados
]

pos = {
    nombre: i
    for i, nombre in enumerate(encabezados)
}

for columna in ["Empleado", "ValidFrom", "ValidTo"]:
    if columna not in pos:
        raise ValueError(
            f"No se encontró la columna '{columna}' "
            "en AbsentismosIT.xlsx."
        )

sick_rows = []

for fila in filas:

    empleado = (
        str(fila[pos["Empleado"]]).strip()
        if fila[pos["Empleado"]] is not None
        else None
    )

    valid_from = (
        str(fila[pos["ValidFrom"]]).strip()
        if fila[pos["ValidFrom"]] is not None
        else None
    )

    valid_to = (
        str(fila[pos["ValidTo"]]).strip()
        if fila[pos["ValidTo"]] is not None
        else None
    )

    if empleado:
        sick_rows.append(
            (
                empleado,
                valid_from,
                valid_to
            )
        )

libro.close()

con.execute("""
    CREATE OR REPLACE TABLE sick_raw (
        empleado_key VARCHAR,
        valid_from VARCHAR,
        valid_to VARCHAR
    )
""")

if sick_rows:
    con.executemany(
        """
        INSERT INTO sick_raw
        VALUES (?, ?, ?)
        """,
        sick_rows
    )

con.execute(f"""
    CREATE OR REPLACE TABLE sick AS
    SELECT DISTINCT

        empleado_key,

        try_strptime(
            nullif(trim(valid_from), ''),
            '%Y%m%d'
        )::DATE AS desde,

        CASE
            WHEN trim(valid_to) = '{FECHA_CENTINELA}'
            THEN NULL
            ELSE try_strptime(
                nullif(trim(valid_to), ''),
                '%Y%m%d'
            )::DATE
        END AS hasta,

        trim(valid_to) = '{FECHA_CENTINELA}'
            AS es_abierto

    FROM sick_raw

    WHERE empleado_key IS NOT NULL

      AND try_strptime(
            nullif(trim(valid_from), ''),
            '%Y%m%d'
          ) IS NOT NULL
""")

invalidos = con.execute("""
    SELECT count(*)
    FROM sick
    WHERE NOT es_abierto
      AND (
          hasta IS NULL
          OR hasta < desde
      )
""").fetchone()[0]

if invalidos:
    raise ValueError(
        f"Se encontraron {invalidos:,} "
        "registros IT inválidos."
    )

total_sick = con.execute(
    "SELECT count(*) FROM sick"
).fetchone()[0]

total_abiertos = con.execute("""
    SELECT count(*)
    FROM sick
    WHERE es_abierto
""").fetchone()[0]

print(f"   IT: {total_sick:,}")
print(f"   Abiertos:  {total_abiertos:,}")


# ============================================================
# 3. RANGO HISTÓRICO
# ============================================================

print("3/9 - Preparando episodios...")

fecha_min_persona = con.execute("""
    SELECT min(inicio_semana)
    FROM persona_base
""").fetchone()[0]

fecha_max_corte = con.execute("""
    SELECT max(inicio_semana) - INTERVAL 1 DAY
    FROM persona_base
""").fetchone()[0]

fecha_min_historico = con.execute("""
    SELECT ?::DATE - INTERVAL 12 MONTH
""", [fecha_min_persona]).fetchone()[0]


con.execute("""
    CREATE OR REPLACE TABLE sick_filtrado AS

    SELECT

        empleado_key,
        desde,
        hasta,
        es_abierto,

        CASE
            WHEN es_abierto
            THEN ?::DATE
            ELSE hasta
        END AS hasta_analisis

    FROM sick

    WHERE desde <= ?::DATE

      AND (
          es_abierto
          OR hasta >= ?::DATE
      )
""", [
    fecha_max_corte,
    fecha_max_corte,
    fecha_min_historico
])


# ============================================================
# 4. GENERACIÓN DIARIA
# ============================================================

print("4/9 - Generando días de IT...")

"""
Se genera una fila por empleado + episodio + día.

La duración del episodio se utiliza exclusivamente para
clasificar cada día como ShortTerm o LongTerm.

Para episodios abiertos, la duración observada hasta la
fecha máxima de corte se utiliza para la clasificación.
"""


con.execute("""
    CREATE OR REPLACE TABLE sick_diario_bruto AS

    SELECT

        s.empleado_key,

        s.desde,

        s.hasta,

        s.es_abierto,

        d.fecha::DATE AS fecha,

        date_trunc(
            'week',
            d.fecha::DATE
        )::DATE AS semana_inicio,

        (
            date_diff(
                'day',
                s.desde,
                s.hasta_analisis
            ) + 1
        ) AS duracion_observada,

        CASE
            WHEN (
                date_diff(
                    'day',
                    s.desde,
                    s.hasta_analisis
                ) + 1
            ) <= 3
            THEN 'short'
            ELSE 'long'
        END AS tipo_episodio

    FROM sick_filtrado s

    CROSS JOIN LATERAL generate_series(

        greatest(
            s.desde,
            ?::DATE
        ),

        s.hasta_analisis,

        INTERVAL 1 DAY

    ) AS d(fecha)

    WHERE s.desde <= s.hasta_analisis
""", [fecha_min_historico])


dias_brutos = con.execute("""
    SELECT count(*)
    FROM sick_diario_bruto
""").fetchone()[0]

print(
    f"   Días calendario generados: {dias_brutos:,}"
)


# ============================================================
# 5. MÁXIMO 5 DÍAS POR EMPLEADO-SEMANA
# ============================================================

print(
    "5/9 - Limitando a un máximo de 5 días por semana..."
)

"""
Para cada empleado y semana:

    1. Se ordenan cronológicamente los días.
    2. Se conservan los primeros 5 días.
    3. Se conserva la clasificación Short/Long del episodio.

Esto garantiza que:

    Total = ShortTerm + LongTerm

y que:

    Total <= 5

No se utiliza TipoJornada.
No se utiliza 5/7.
No se utilizan decimales.
"""


con.execute("""
    CREATE OR REPLACE TABLE sick_diario AS

    WITH ordenado AS (

        SELECT

            *,

            row_number() OVER (

                PARTITION BY
                    empleado_key,
                    semana_inicio

                ORDER BY
                    fecha,
                    desde

            ) AS numero_dia_semana

        FROM sick_diario_bruto
    )

    SELECT

        empleado_key,
        fecha,
        semana_inicio,
        tipo_episodio

    FROM ordenado

    WHERE numero_dia_semana <= 5
""")


# ============================================================
# CONSOLIDACIÓN SEMANAL
# ============================================================

con.execute("""
    CREATE OR REPLACE TABLE sick_semanal AS

    SELECT

        empleado_key,
        semana_inicio,

        CAST(
            count(*) AS BIGINT
        ) AS total_dia,

        CAST(
            sum(
                CASE
                    WHEN tipo_episodio = 'short'
                    THEN 1
                    ELSE 0
                END
            ) AS BIGINT
        ) AS short_dia,

        CAST(
            sum(
                CASE
                    WHEN tipo_episodio = 'long'
                    THEN 1
                    ELSE 0
                END
            ) AS BIGINT
        ) AS long_dia

    FROM sick_diario

    GROUP BY
        empleado_key,
        semana_inicio
""")


# ============================================================
# CONTROLES SEMANALES
# ============================================================

inconsistencias = con.execute("""
    SELECT count(*)
    FROM sick_semanal
    WHERE total_dia
        != short_dia + long_dia
""").fetchone()[0]

if inconsistencias:
    raise ValueError(
        f"Se encontraron {inconsistencias:,} "
        "semanas donde Total != ShortTerm + LongTerm."
    )


excesos = con.execute("""
    SELECT count(*)
    FROM sick_semanal
    WHERE total_dia > 5
""").fetchone()[0]

if excesos:
    raise ValueError(
        f"Se encontraron {excesos:,} "
        "semanas con más de 5 días."
    )


decimales = con.execute("""
    SELECT count(*)
    FROM sick_semanal
    WHERE
        total_dia != floor(total_dia)
        OR short_dia != floor(short_dia)
        OR long_dia != floor(long_dia)
""").fetchone()[0]

if decimales:
    raise ValueError(
        f"Se encontraron {decimales:,} "
        "valores decimales."
    )


total_semanas_sick = con.execute("""
    SELECT count(*)
    FROM sick_semanal
""").fetchone()[0]

print(
    f"   Empleado-semana con IT: "
    f"{total_semanas_sick:,}"
)


# ============================================================
# 6. ACUMULADOS SEMANALES
# ============================================================

print(
    "6/9 - Calculando acumulados..."
)


"""
Los acumulados se calculan a nivel semanal.

Las ventanas posteriores se calculan sobre estas semanas,
pero se mantienen los límites temporales originales.
"""


con.execute("""
    CREATE OR REPLACE TABLE sick_acumulado AS

    SELECT

        empleado_key,
        semana_inicio,

        total_dia,
        short_dia,
        long_dia,

        SUM(total_dia) OVER (
            PARTITION BY empleado_key
            ORDER BY semana_inicio
            ROWS BETWEEN
                UNBOUNDED PRECEDING
                AND CURRENT ROW
        ) AS cum_total,

        SUM(short_dia) OVER (
            PARTITION BY empleado_key
            ORDER BY semana_inicio
            ROWS BETWEEN
                UNBOUNDED PRECEDING
                AND CURRENT ROW
        ) AS cum_short,

        SUM(long_dia) OVER (
            PARTITION BY empleado_key
            ORDER BY semana_inicio
            ROWS BETWEEN
                UNBOUNDED PRECEDING
                AND CURRENT ROW
        ) AS cum_long

    FROM sick_semanal
""")


# ============================================================
# PREPARACIÓN DE VENTANAS
# ============================================================

con.execute("""
    CREATE OR REPLACE TABLE ventanas AS

    SELECT

        *,

        inicio_semana
            - INTERVAL 1 DAY
            AS fecha_corte,

        inicio_semana
            - INTERVAL 7 DAY
            AS inicio_1w,

        inicio_semana
            - INTERVAL 14 DAY
            AS inicio_2w,

        inicio_semana
            - INTERVAL 1 MONTH
            AS inicio_1m,

        inicio_semana
            - INTERVAL 3 MONTH
            AS inicio_3m,

        inicio_semana
            - INTERVAL 6 MONTH
            AS inicio_6m,

        inicio_semana
            - INTERVAL 12 MONTH
            AS inicio_12m

    FROM persona_base
""")


def generar_ventana(
    nombre,
    columna_inicio
):
    """
    Calcula una ventana histórica mediante acumulados semanales.

    Se incluyen las semanas cuyo inicio se encuentra:

        >= inicio de ventana

    y

        <= semana anterior a la semana objetivo.

    Para ventanas mensuales se utiliza la semana cuyo lunes
    está dentro del intervalo temporal.
    """

    con.execute(f"""

        CREATE OR REPLACE TABLE ventana_{nombre} AS

        SELECT

            p."ID_Semana"
                AS id_semana,

            p.empleado_key,

            CAST(
                COALESCE(
                    SUM(
                        CASE
                            WHEN a.semana_inicio
                                >= p.{columna_inicio}
                            AND a.semana_inicio
                                <= p.fecha_corte
                            THEN a.total_dia
                            ELSE 0
                        END
                    ),
                    0
                )
                AS BIGINT
            ) AS total,

            CAST(
                COALESCE(
                    SUM(
                        CASE
                            WHEN a.semana_inicio
                                >= p.{columna_inicio}
                            AND a.semana_inicio
                                <= p.fecha_corte
                            THEN a.short_dia
                            ELSE 0
                        END
                    ),
                    0
                )
                AS BIGINT
            ) AS short,

            CAST(
                COALESCE(
                    SUM(
                        CASE
                            WHEN a.semana_inicio
                                >= p.{columna_inicio}
                            AND a.semana_inicio
                                <= p.fecha_corte
                            THEN a.long_dia
                            ELSE 0
                        END
                    ),
                    0
                )
                AS BIGINT
            ) AS long

        FROM ventanas p

        LEFT JOIN sick_acumulado a

            ON p.empleado_key
                = a.empleado_key

        GROUP BY

            p."ID_Semana",
            p.empleado_key,
            p.{columna_inicio},
            p.fecha_corte

    """)


generar_ventana(
    "1Week",
    "inicio_1w"
)

generar_ventana(
    "2Weeks",
    "inicio_2w"
)

generar_ventana(
    "1Month",
    "inicio_1m"
)

generar_ventana(
    "3Months",
    "inicio_3m"
)

generar_ventana(
    "6Months",
    "inicio_6m"
)

generar_ventana(
    "12Months",
    "inicio_12m"
)


# ============================================================
# 7. GENERACIÓN DE TARGETS FUTUROS
# ============================================================

print(
    "7/9 - Generando Targets futuros..."
)

"""
Targets para cada empleado y semana de referencia W:

    IT_Futura_W1 = días IT en W+1
    IT_Futura_W2 = días IT en W+2
    IT_Futura_W3 = días IT en W+3
    IT_Futura_W4 = días IT en W+4

Los Targets se calculan directamente desde AbsentismosIT.xlsx.
NO se utilizan las variables Previous... para construirlos.

La tabla sick_semanal utilizada para el histórico está limitada
al corte de Persona_Semana. Por eso los Targets se construyen
con una tabla diaria/semanal independiente que puede extenderse
hasta las semanas futuras necesarias.

Cada semana futura conserva la misma regla de negocio utilizada
en el histórico: máximo 5 días computables por empleado-semana.
"""

# Rango de semanas futuras necesario para todos los Targets.
fecha_min_target = con.execute("""
    SELECT min(inicio_semana) + INTERVAL 7 DAY
    FROM persona_base
""").fetchone()[0]

fecha_max_target = con.execute("""
    SELECT max(inicio_semana) + INTERVAL 34 DAY
    FROM persona_base
""").fetchone()[0]

# Última fecha cerrada disponible en la fuente de IT.
fecha_max_it = con.execute("""
    SELECT max(hasta)
    FROM sick
    WHERE hasta IS NOT NULL
""").fetchone()[0]

print(
    f"   Primer día necesario para Targets: {fecha_min_target}"
)
print(
    f"   Último día necesario para Targets: {fecha_max_target}"
)
print(
    f"   Última fecha cerrada disponible en IT: {fecha_max_it}"
)

# Generación diaria EXCLUSIVA para Targets.
# No modifica sick_diario ni sick_semanal históricos.
con.execute("""
    CREATE OR REPLACE TABLE sick_diario_target_bruto AS

    SELECT

        s.empleado_key,
        s.desde,
        s.hasta,
        s.es_abierto,
        d.fecha::DATE AS fecha,
        date_trunc(
            'week',
            d.fecha::DATE
        )::DATE AS semana_inicio

    FROM sick s

    CROSS JOIN LATERAL generate_series(

        greatest(
            s.desde,
            ?::DATE
        ),

        least(
            CASE
                WHEN s.es_abierto
                THEN ?::DATE
                ELSE s.hasta
            END,
            ?::DATE
        ),

        INTERVAL 1 DAY

    ) AS d(fecha)

    WHERE s.desde <= ?::DATE

      AND (
          s.es_abierto
          OR s.hasta >= ?::DATE
      )
""", [
    fecha_min_target,
    fecha_max_target,
    fecha_max_target,
    fecha_max_target,
    fecha_min_target
])

# Máximo 5 días por empleado-semana, igual que en el histórico.
con.execute("""
    CREATE OR REPLACE TABLE sick_diario_target AS

    WITH ordenado AS (

        SELECT
            *,
            row_number() OVER (
                PARTITION BY
                    empleado_key,
                    semana_inicio
                ORDER BY
                    fecha,
                    desde
            ) AS numero_dia_semana

        FROM sick_diario_target_bruto
    )

    SELECT
        empleado_key,
        fecha,
        semana_inicio

    FROM ordenado

    WHERE numero_dia_semana <= 5
""")

con.execute("""
    CREATE OR REPLACE TABLE sick_semanal_target AS

    SELECT
        empleado_key,
        semana_inicio,
        CAST(count(*) AS BIGINT) AS total_dia

    FROM sick_diario_target

    GROUP BY
        empleado_key,
        semana_inicio
""")

# Targets por fila de Persona_Semana.
# Si una semana está cubierta por la fuente de IT y no hay IT,
# el valor es 0. Si queda fuera de la cobertura de la fuente,
# queda NULL.
con.execute("""
    CREATE OR REPLACE TABLE targets_futuros AS

    SELECT

        p."ID_Semana" AS id_semana,
        p.empleado_key,

        CASE
            WHEN (
                p.inicio_semana + INTERVAL 13 DAY
            ) <= ?::DATE
            THEN COALESCE(w1.total_dia, 0)
            ELSE NULL
        END AS IT_Futura_W1,

        CASE
            WHEN (
                p.inicio_semana + INTERVAL 20 DAY
            ) <= ?::DATE
            THEN COALESCE(w2.total_dia, 0)
            ELSE NULL
        END AS IT_Futura_W2,

        CASE
            WHEN (
                p.inicio_semana + INTERVAL 27 DAY
            ) <= ?::DATE
            THEN COALESCE(w3.total_dia, 0)
            ELSE NULL
        END AS IT_Futura_W3,

        CASE
            WHEN (
                p.inicio_semana + INTERVAL 34 DAY
            ) <= ?::DATE
            THEN COALESCE(w4.total_dia, 0)
            ELSE NULL
        END AS IT_Futura_W4

    FROM persona_base p

    LEFT JOIN sick_semanal_target w1
        ON p.empleado_key = w1.empleado_key
       AND w1.semana_inicio =
            p.inicio_semana + INTERVAL 7 DAY

    LEFT JOIN sick_semanal_target w2
        ON p.empleado_key = w2.empleado_key
       AND w2.semana_inicio =
            p.inicio_semana + INTERVAL 14 DAY

    LEFT JOIN sick_semanal_target w3
        ON p.empleado_key = w3.empleado_key
       AND w3.semana_inicio =
            p.inicio_semana + INTERVAL 21 DAY

    LEFT JOIN sick_semanal_target w4
        ON p.empleado_key = w4.empleado_key
       AND w4.semana_inicio =
            p.inicio_semana + INTERVAL 28 DAY
""", [
    fecha_max_it,
    fecha_max_it,
    fecha_max_it,
    fecha_max_it
])

targets_w1_disponible = con.execute("""
    SELECT count(*)
    FROM targets_futuros
    WHERE IT_Futura_W1 IS NOT NULL
""").fetchone()[0]

targets_4_disponibles = con.execute("""
    SELECT count(*)
    FROM targets_futuros
    WHERE IT_Futura_W1 IS NOT NULL
      AND IT_Futura_W2 IS NOT NULL
      AND IT_Futura_W3 IS NOT NULL
      AND IT_Futura_W4 IS NOT NULL
""").fetchone()[0]

print(
    f"   Filas con W+1 disponible:        {targets_w1_disponible:,}"
)
print(
    f"   Filas con W+1..W+4 completos:   {targets_4_disponibles:,}"
)


# ============================================================
# 8. INTEGRACIÓN FINAL
# ============================================================

print(
    "8/9 - Integrando tabla final..."
)


con.execute("""
    CREATE OR REPLACE TABLE resultado AS

    SELECT

        p.*,

        CASE

            WHEN EXISTS (

                SELECT 1

                FROM sick s

                WHERE s.empleado_key
                    = p.empleado_key

                  AND s.es_abierto

                  AND s.desde <= (
                      p.inicio_semana
                      - INTERVAL 1 DAY
                  )

            )

            THEN 1

            ELSE 0

        END AS ITAbierto

    FROM persona_base p
""")


# Añadir Targets futuros
con.execute("""
    CREATE OR REPLACE TABLE resultado AS

    SELECT
        r.*,
        t.IT_Futura_W1,
        t.IT_Futura_W2,
        t.IT_Futura_W3,
        t.IT_Futura_W4

    FROM resultado r

    LEFT JOIN targets_futuros t
        ON r."ID_Semana" = t.id_semana
       AND r.empleado_key = t.empleado_key
""")


ventanas_config = [

    (
        "1Week",
        "Previous1Week"
    ),

    (
        "2Weeks",
        "Previous2Weeks"
    ),

    (
        "1Month",
        "Previous1Month"
    ),

    (
        "3Months",
        "Previous3Months"
    ),

    (
        "6Months",
        "Previous6Months"
    ),

    (
        "12Months",
        "Previous12Months"
    )

]


for tabla, prefijo in ventanas_config:

    con.execute(f"""

        CREATE OR REPLACE TABLE resultado AS

        SELECT

            r.*,

            CAST(
                v.total AS BIGINT
            ) AS "{prefijo}IT",

            CAST(
                v.short AS BIGINT
            ) AS "{prefijo}ITShortTerm",

            CAST(
                v.long AS BIGINT
            ) AS "{prefijo}ITLongTerm"

        FROM resultado r

        INNER JOIN ventana_{tabla} v

            ON r."ID_Semana"
                = v.id_semana

           AND r.empleado_key
                = v.empleado_key

    """)


# ============================================================
# ELEGIBILIDAD
# ============================================================

con.execute("""
    CREATE OR REPLACE TABLE resultado AS

    SELECT

        *,

        CASE

            WHEN ITAbierto = 1
            THEN 0

            ELSE 1

        END AS ElegiblePrediccion

    FROM resultado
""")


# ============================================================
# CONTROLES FINALES
# ============================================================

print(
    "   Ejecutando controles finales..."
)


for prefijo in [

    "Previous1Week",
    "Previous2Weeks",
    "Previous1Month",
    "Previous3Months",
    "Previous6Months",
    "Previous12Months"

]:

    total_col = f'"{prefijo}IT"'
    short_col = f'"{prefijo}ITShortTerm"'
    long_col = f'"{prefijo}ITLongTerm"'

    inconsistencias = con.execute(f"""

        SELECT count(*)

        FROM resultado

        WHERE

            {total_col}
            !=
            {short_col}
            +
            {long_col}

    """).fetchone()[0]

    if inconsistencias:

        raise ValueError(

            f"La ventana {prefijo} tiene "
            f"{inconsistencias:,} filas donde "
            "Total != ShortTerm + LongTerm."

        )


    decimales = con.execute(f"""

        SELECT count(*)

        FROM resultado

        WHERE

            {total_col}
                != floor({total_col})

            OR

            {short_col}
                != floor({short_col})

            OR

            {long_col}
                != floor({long_col})

    """).fetchone()[0]

    if decimales:

        raise ValueError(

            f"La ventana {prefijo} contiene "
            f"{decimales:,} valores decimales."

        )


# ============================================================
# CONTROLES DE TARGETS
# ============================================================

for target in [
    "IT_Futura_W1",
    "IT_Futura_W2",
    "IT_Futura_W3",
    "IT_Futura_W4"
]:

    negativos = con.execute(f"""
        SELECT count(*)
        FROM resultado
        WHERE "{target}" IS NOT NULL
          AND "{target}" < 0
    """).fetchone()[0]

    if negativos:
        raise ValueError(
            f"{target} contiene {negativos:,} valores negativos."
        )

    excesos_target = con.execute(f"""
        SELECT count(*)
        FROM resultado
        WHERE "{target}" IS NOT NULL
          AND "{target}" > 5
    """).fetchone()[0]

    if excesos_target:
        raise ValueError(
            f"{target} contiene {excesos_target:,} filas con más de 5 días."
        )


# ============================================================
# CONTROL ESPECÍFICO EMPLEADO 23109
# ============================================================

print()
print(
    "Control específico empleado 23109..."
)


control_23109 = con.execute("""

    SELECT

        "CodEmpleado",
        "ID_Semana",
        "FechaInicio",
        "FechaFin",

        "Previous1WeekIT",
        "Previous1WeekITShortTerm",
        "Previous1WeekITLongTerm",

        "Previous2WeeksIT",
        "Previous2WeeksITShortTerm",
        "Previous2WeeksITLongTerm",

        "Previous1MonthIT",
        "Previous1MonthITShortTerm",
        "Previous1MonthITLongTerm"

    FROM resultado

    WHERE trim(
        cast("CodEmpleado" AS VARCHAR)
    ) = '23109'

      AND "ID_Semana"
          BETWEEN 20 AND 22

    ORDER BY
        "ID_Semana"

""").fetchall()


if control_23109:

    print()

    for fila in control_23109:

        print(

            f"   Semana {fila[1]} | "
            f"Prev1={fila[4]} "
            f"(Short={fila[5]}, Long={fila[6]}) | "
            f"Prev2={fila[7]} "
            f"(Short={fila[8]}, Long={fila[9]}) | "
            f"Prev1M={fila[10]} "
            f"(Short={fila[11]}, Long={fila[12]})"

        )

else:

    print(
        "   No se encontró el empleado 23109."
    )


# ============================================================
# EXPORTACIÓN
# ============================================================

print(
    "9/9 - Exportando resultados..."
)


columnas = [

    fila[0]

    for fila in con.execute(
        "DESCRIBE resultado"
    ).fetchall()

]


columnas_salida = [

    c

    for c in columnas

    if c not in {

        "empleado_key",
        "inicio_semana",
        "fin_semana"

    }

]


select_salida = ", ".join(

    '"' + c.replace(
        '"',
        '""'
    ) + '"'

    for c in columnas_salida

)


con.execute(f"""

    COPY (

        SELECT

            {select_salida}

        FROM resultado

        ORDER BY

            "ID_Semana",
            "CodEmpleado"

    )

    TO '{ruta_sql(ARCHIVO_SALIDA)}'

    (

        HEADER,
        DELIMITER ';',
        QUOTE '"',
        ESCAPE '"'

    )

""")


# ============================================================
# CONTROLES GENERALES
# ============================================================

filas_finales = con.execute("""
    SELECT count(*)
    FROM resultado
""").fetchone()[0]


duplicados_finales = con.execute("""

    SELECT COUNT(*)

    FROM (

        SELECT

            "CodEmpleado",
            "ID_Semana"

        FROM resultado

        GROUP BY

            "CodEmpleado",
            "ID_Semana"

        HAVING COUNT(*) > 1

    )

""").fetchone()[0]


filas_no_elegibles = con.execute("""

    SELECT count(*)

    FROM resultado

    WHERE ElegiblePrediccion = 0

""").fetchone()[0]


filas_elegibles = con.execute("""

    SELECT count(*)

    FROM resultado

    WHERE ElegiblePrediccion = 1

""").fetchone()[0]


# ============================================================
# RESUMEN
# ============================================================

resumen = {

    "estado": "OK",

    "filas_persona_semana":
        int(total_persona),

    "filas_tabla_entrenamiento":
        int(filas_finales),

    "empleados_persona_semana":
        int(
            con.execute("""
                SELECT COUNT(
                    DISTINCT "CodEmpleado"
                )
                FROM resultado
            """).fetchone()[0]
        ),

    "registros_it":
        int(total_sick),

    "registros_it_abiertos":
        int(total_abiertos),

    "dias_calendario_generados":
        int(dias_brutos),

    "combinaciones_empleado_semana_it":
        int(total_semanas_sick),

    "variables_historicas":
        18,

    "variables_target":
        4,

    "targets":
        [
            "IT_Futura_W1",
            "IT_Futura_W2",
            "IT_Futura_W3",
            "IT_Futura_W4"
        ],

    "filas_target_w1_disponible":
        int(targets_w1_disponible),

    "filas_target_4_semanas_completo":
        int(targets_4_disponibles),

    "fecha_max_it_cerrada":
        str(fecha_max_it) if fecha_max_it is not None else None,

    "variables_por_ventana":
        3,

    "ventanas":
        [
            "1 semana",
            "2 semanas",
            "1 mes",
            "3 meses",
            "6 meses",
            "12 meses"
        ],

    "filas_elegibles_prediccion":
        int(filas_elegibles),

    "filas_no_elegibles_baja_abierta":
        int(filas_no_elegibles),

    "duplicados_empleado_semana":
        int(duplicados_finales),

    "valores_decimales_it":
        0,

    "inconsistencias_total_short_long":
        0,

    "maximo_dias_it_semana":
        5,

    "regla_dias":
        (
            "Se generan días calendario por episodio. "
            "Para cada empleado-semana se conservan como "
            "máximo 5 días, ordenados cronológicamente. "
            "No se utiliza TipoJornada ni ponderación 5/7."
        ),

    "regla_short_long":
        {
            "shortterm":
                "<= 3 días calendario de duración observada "
                "del episodio.",

            "longterm":
                "> 3 días calendario de duración observada "
                "del episodio."
        },

    "regla_consistencia":
        "Total = ShortTerm + LongTerm.",

    "regla_baja_abierta":
        (
            "Las bajas abiertas se conservan para histórico "
            "y hacen que el empleado no sea elegible mientras "
            "permanecen abiertas."
        ),

    "tipo_jornada_utilizado":
        False,

    "ponderacion_5_7_utilizada":
        False,

    "variables_por_dia_semana":
        False

}


ARCHIVO_RESUMEN.write_text(

    json.dumps(
        resumen,
        ensure_ascii=False,
        indent=2
    ),

    encoding="utf-8"

)


con.close()


# ============================================================
# FIN
# ============================================================

print()
print("=" * 70)
print("TABLA DE ENTRENAMIENTO GENERADA CORRECTAMENTE")
print("=" * 70)

print(
    f"Filas de entrada:             "
    f"{total_persona:,}"
)

print(
    f"Filas de salida:              "
    f"{filas_finales:,}"
)

print(
    "Variables históricas:         18"
)

print(
    "Variables Target:             4"
)

print(
    "Targets:                      W+1, W+2, W+3, W+4"
)

print(
    f"Filas con Target W+1:         {targets_w1_disponible:,}"
)

print(
    f"Filas con Target W+1..W+4:   {targets_4_disponibles:,}"
)

print(
    f"IT abiertos:                  {total_abiertos:,}"
)

print(
    f"Empleado-semana IT:           {total_semanas_sick:,}"
)

print(
    f"Filas elegibles:              {filas_elegibles:,}"
)

print(
    f"Filas no elegibles:           {filas_no_elegibles:,}"
)

print(
    f"Duplicados empleado-semana:   {duplicados_finales:,}"
)

print(
    "Valores decimales IT:         0"
)

print(
    "Total = Short + Long:         OK"
)

print(
    "Máximo semanal:               5"
)

print(
    f"CSV:                          {ARCHIVO_SALIDA}"
)

print(
    f"JSON:                         {ARCHIVO_RESUMEN}"
)

print("=" * 70)
