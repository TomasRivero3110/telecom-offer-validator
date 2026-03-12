# -*- coding: utf-8 -*-
from __future__ import print_function

import cx_Oracle

# -------------------------------------------------------------
# Conexión a la BD (misma que usás en los otros scripts)
# -------------------------------------------------------------

connection = cx_Oracle.connect(
    "DB_USER",
    "DB_PASSWORD",
    "DB_USER_SERVER",
    encoding="UTF-8",
    nencoding="UTF-8"
)

listaBonosParent = []


# -------------------------------------------------------------
# FUNCIÓN PRINCIPAL: BonosPorParent(activity_id)
# -------------------------------------------------------------

def BonosPorParent(activity_id):
    """
    Devuelve una lista donde cada fila contiene:
    [category, parent_offeringID, parent_offering_name, bonos_list]

    bonos_list = todos los bonos hijo del parent, separados por coma.

    SOLO:
      - Paquetes (parent) que están en el CR (nc_cmt_ct_modifs sobre el parent)
      - Bonos que están REALMENTE vinculados como hijos en R_PIM_POFF_CDL_RLT
      - No incluye bonos sueltos ni bonos no vinculados al paquete.
    """
    global connection, listaBonosParent

    listaBonosParent = []
    cursor = connection.cursor()

    sql = """
    WITH base AS (
        SELECT DISTINCT
            cat.name       AS category,
            parent.PRD_OFF AS parent_offeringID,
            parent.name    AS parent_offering_name,
            bono.PRD_OFF   AS offeringID,
            bono.name      AS offering_name
        FROM prd_app_6800.R_PIM_PRD_OFF parent

        -- CR asociado SOLO al PAQUETE
        JOIN prd_app_6800.nc_cmt_ct_modifs cr_parent
             ON cr_parent.ci_id = parent.PRD_OFF

        -- Relación paquete -> bono
        JOIN prd_app_6800.R_PIM_POFF_CDL_RLT rel_parent
             ON rel_parent.PARENT = parent.PRD_OFF
        JOIN prd_app_6800.R_PIM_PRD_OFF bono
             ON bono.PRD_OFF = rel_parent.CHILD

        -- Categoría del paquete (opcional, para mostrar category)
        LEFT JOIN prd_app_6800.R_PIM_PCHLD_OF_RLHP rel_cat
               ON parent.PRD_OFF = rel_cat.CHILD
        LEFT JOIN prd_app_6800.R_PIM_OFFR_CAT cat
               ON cat.offr_cat = rel_cat.PARENT

        WHERE cr_parent.activity_id = :activity_id
          AND parent.name IS NOT NULL
          AND bono.name LIKE 'Bono%'   -- SOLO BONOS hijos del paquete
    ),

    agg AS (
        SELECT
            MIN(category) AS category,
            parent_offeringID,
            parent_offering_name,
            LISTAGG(offering_name, ', ')
                WITHIN GROUP (ORDER BY offering_name) AS bonos_list
        FROM base
        GROUP BY
            parent_offeringID,
            parent_offering_name
    ),

    dedup AS (
        SELECT
            category,
            parent_offeringID,
            parent_offering_name,
            bonos_list,
            ROW_NUMBER() OVER (
                PARTITION BY parent_offering_name
                ORDER BY parent_offeringID DESC   -- nos quedamos con el ID más nuevo
            ) AS rn
        FROM agg
    )

    SELECT
        category,
        parent_offeringID,
        parent_offering_name,
        bonos_list
    FROM dedup
    WHERE rn = 1
    ORDER BY
        category,
        parent_offering_name
    """


    cursor.execute(sql, activity_id=activity_id)

    for row in cursor:
        listaBonosParent.append(list(row))

    cursor.close()
    return listaBonosParent
