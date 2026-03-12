# -*- coding: utf-8 -*-
from __future__ import print_function
import cx_Oracle
from decimal import InvalidOperation  # Ya no se usa; podés borrarlo si querés

connection = cx_Oracle.connect(
    "app_catalogo",
    "C4t4logo_2020",
    "10.24.135.33:1521/NTSTTOMS",
    encoding="UTF-8",
    nencoding="UTF-8"
)

listaBonosPaquete = []


def normalizar_amount(valor):
    """
    Normaliza el amount para que, si viene en notación científica
    (por ejemplo '1.073741824E+10', '6E+2', '5.36870912E+9', etc.),
    lo devuelva como número plano en formato string: '10737418240', '600', etc.

    Si no se puede convertir a número (texto raro), lo devuelve tal cual.
    Acepta None, LOBs y tipos básicos.
    """
    if valor is None:
        return None

    s = str(valor).strip()
    try:
        num = float(s)          # entiende notación científica muy permisivamente
    except ValueError:
        # No es un número, devolvemos el texto como está
        return s

    # Si es entero exacto, lo dejamos sin decimales
    if num.is_integer():
        return str(int(num))
    else:
        # Si tuviera decimales reales, los dejamos sin ceros sobrantes
        txt = "{:.15f}".format(num).rstrip('0').rstrip('.')
        return txt


def BonosPorPaqueteDetallado(activity_id):
    """
    Devuelve TODOS los bonos asociados a paquetes en el CR,
    sin repetir combinaciones iguales.

    Cada fila incluye:

    0  - parent_offeringID
    1  - parent_offering_name
    2  - offeringID (ID del bono)
    3  - offering_name (nombre del bono)
    4  - validity_period (AA Validity Period)
    5  - amount (AA Amount, normalizado como string sin E+10)
    6  - bonus_code (AA Package Object ID)
    7  - bonus_description (AA Bonus Description / Description)
    8  - child_min        (CHILD_MIN de la relación paquete→bono)
    9  - child_max        (CHILD_MAX de la relación paquete→bono)
    10 - DEFAULT_BEHAVIOR (DEFAULT_BEHAVIOR de la relación)
    11 - hide_component   (HIDE_CMPN de la relación)
    12 - object_type      (característica 'AA Object Type' del bono)

    - Solo se consideran PAQUETES cuyo CI está en el CR.
    - Para cada paquete, se traen todos los bonos realmente vinculados,
      aunque el bono no tenga entrada propia en el CR.
    """

    global connection, listaBonosPaquete
    listaBonosPaquete = []
    cursor = connection.cursor()

    sql = """
    WITH bonos_raw AS (
        SELECT
            parent.PRD_OFF    AS parent_offeringID,
            parent.name       AS parent_offering_name,

            po.PRD_OFF        AS offeringID,
            po.name           AS offering_name,

            /* Período del bono */
            MAX(
                CASE
                    WHEN chv.name = 'AA Validity Period' THEN
                        COALESCE(
                            chv.DEFAULT_VALUE_TEXT,
                            chv.DEFAULT_VALUE_NUMB,
                            chv.DEFAULT_VALUE_DECIM,
                            TO_CHAR(chv.DEF_LIST_VALUE_NUMBER),
                            TO_CHAR(chv.DEF_LIST_VALUE_DECIMAL),
                            TO_CHAR(chv.DEF_LIST_VALUE),
                            TO_CHAR(chv.DEF_LIST_VALUE_TEXT),
                            TO_CHAR(chv.DEFAULT_VALUE_DATE,   'YYYY-MM-DD'),
                            TO_CHAR(chv.DEF_LIST_VALUE_DATE, 'YYYY-MM-DD')
                        )
                END
            ) AS validity_period,

            /* AA Amount */
            MAX(
                CASE
                    WHEN chv.name = 'AA Amount' THEN
                        COALESCE(
                            chv.DEFAULT_VALUE_TEXT,
                            chv.DEFAULT_VALUE_NUMB,
                            chv.DEFAULT_VALUE_DECIM,
                            TO_CHAR(chv.DEF_LIST_VALUE_NUMBER),
                            TO_CHAR(chv.DEF_LIST_VALUE_DECIMAL),
                            TO_CHAR(chv.DEF_LIST_VALUE),
                            TO_CHAR(chv.DEF_LIST_VALUE_TEXT)
                        )
                END
            ) AS amount,

            /* AA Package Object ID → "bonus code" */
            MAX(
                CASE
                    WHEN chv.name = 'AA Package Object ID' THEN
                        COALESCE(
                            chv.DEFAULT_VALUE_TEXT,
                            chv.DEFAULT_VALUE_NUMB,
                            chv.DEFAULT_VALUE_DECIM,
                            TO_CHAR(chv.DEF_LIST_VALUE_NUMBER),
                            TO_CHAR(chv.DEF_LIST_VALUE_DECIMAL),
                            TO_CHAR(chv.DEF_LIST_VALUE),
                            TO_CHAR(chv.DEF_LIST_VALUE_TEXT)
                        )
                END
            ) AS bonus_code,

            /* Bonus Description */
            MAX(
                CASE
                    WHEN chv.name IN ('AA Bonus Description', 'Bonus Description', 'Description') THEN
                        COALESCE(
                            chv.DEFAULT_VALUE_TEXT,
                            chv.DEFAULT_VALUE_NUMB,
                            chv.DEFAULT_VALUE_DECIM,
                            TO_CHAR(chv.DEF_LIST_VALUE_NUMBER),
                            TO_CHAR(chv.DEF_LIST_VALUE_DECIMAL),
                            TO_CHAR(chv.DEF_LIST_VALUE),
                            TO_CHAR(chv.DEF_LIST_VALUE_TEXT)
                        )
                END
            ) AS bonus_description,

            /* AA Object Type - nueva característica */
            MAX(
                CASE
                    WHEN chv.name = 'AA Object Type' THEN
                        COALESCE(
                            chv.DEFAULT_VALUE_TEXT,
                            chv.DEFAULT_VALUE_NUMB,
                            chv.DEFAULT_VALUE_DECIM,
                            TO_CHAR(chv.DEF_LIST_VALUE_NUMBER),
                            TO_CHAR(chv.DEF_LIST_VALUE_DECIMAL),
                            TO_CHAR(chv.DEF_LIST_VALUE),
                            TO_CHAR(chv.DEF_LIST_VALUE_TEXT)
                        )
                END
            ) AS object_type,

            /* Datos de la relación paquete → bono */
            rel_parent.CHILD_MIN        AS child_min,
            rel_parent.CHILD_MAX        AS child_max,
            rel_parent.DEFAULT_BEHAVIOR,
            rel_parent.HIDE_CMPN        AS hide_component

        FROM prd_app_6800.R_PIM_PRD_OFF po

        /* Relación paquete → bono */
        LEFT JOIN prd_app_6800.R_PIM_POFF_CDL_RLT rel_parent
            ON po.PRD_OFF = rel_parent.CHILD

        LEFT JOIN prd_app_6800.R_PIM_PRD_OFF parent
            ON parent.PRD_OFF = rel_parent.PARENT

        /* Características del bono */
        LEFT JOIN prd_app_6800.R_PIM_OFFR_CHR_INV chv
            ON po.PRD_OFF = chv.PARENT_ID

        /* CR asociado al PAQUETE (no al bono) */
        LEFT JOIN prd_app_6800.nc_cmt_ct_modifs cr_parent
            ON cr_parent.ci_id = parent.PRD_OFF

        WHERE cr_parent.activity_id = :activity_id
          AND parent.name IS NOT NULL
          AND (
                parent.name LIKE 'Paquete%'   -- paquetes "Paquete Eventual/Recurrente ..."
             OR parent.name LIKE 'Paq.%'      -- paquetes "Paq. Recurrente PrePack ..."
          )
          AND po.name LIKE 'Bono%'            -- solo bonos hijos

        GROUP BY
            parent.PRD_OFF,
            parent.name,
            po.PRD_OFF,
            po.name,
            rel_parent.CHILD_MIN,
            rel_parent.CHILD_MAX,
            rel_parent.DEFAULT_BEHAVIOR,
            rel_parent.HIDE_CMPN
    )

    SELECT DISTINCT
        parent_offeringID,
        parent_offering_name,
        offeringID,
        offering_name,
        validity_period,
        amount,
        bonus_code,
        bonus_description,
        child_min,
        child_max,
        DEFAULT_BEHAVIOR,
        hide_component,
        object_type

    FROM bonos_raw
    ORDER BY
        parent_offering_name,
        offering_name,
        validity_period,
        amount,
        bonus_code
    """

    cursor.execute(sql, activity_id=activity_id)

    for row in cursor:
        fila = list(row)
        # Índice 5 = amount (0..12)
        fila[5] = normalizar_amount(fila[5])
        listaBonosPaquete.append(fila)

    cursor.close()
    return listaBonosPaquete
