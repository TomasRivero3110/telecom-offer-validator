# -*- coding: utf-8 -*-
from __future__ import print_function

import cx_Oracle
import re
from decimal import Decimal, InvalidOperation

# Conexión a la BD (mismo formato que usás en otros scripts)
connection = cx_Oracle.connect(
    "app_catalogo",
    "C4t4logo_2020",
    "10.24.135.33:1521/NTSTTOMS",
    encoding='UTF-8',
    nencoding='UTF-8'
)

listaPaquetes = []

# ----------------------------------------------------------------------
# Función auxiliar: normalizar AA Amount dentro de CHARACTERISTICS_LIST
# ----------------------------------------------------------------------

# Busca cosas del estilo: "AA Amount: 1.073741824E+10" o "AA Amount: 6E+5"
pat_aa_amount = re.compile(r'(AA Amount:\s*)([^,\s]+)')

def normalizar_aa_amount(cadena):
    """
    Busca 'AA Amount: valor' dentro del string y, si puede interpretar
    'valor' como número (incluyendo notación científica), lo reescribe
    sin E+10. Si no puede convertir, deja el texto como está.
    Acepta None y LOBs de cx_Oracle.
    """
    if cadena is None:
        return None

    # Si viene como LOB u otro tipo, lo pasamos a string
    cadena = str(cadena)

    def _reemplazo(match):
        prefijo, valor = match.groups()
        try:
            num = Decimal(valor)        # entiende 1.073741824E+10, 6E+5, etc.
            plano = format(num, 'f')    # 10737418240.0000, 600000.0, etc.

            if '.' in plano:
                plano = plano.rstrip('0').rstrip('.')

            return prefijo + plano
        except InvalidOperation:
            return match.group(0)

    return pat_aa_amount.sub(_reemplazo, cadena)



# ----------------------------------------------------------------------
# Función principal: Paquetes(activity_id)
# ----------------------------------------------------------------------

def Paquetes(activity_id):
    """
    Devuelve la lista de paquetes para el activity_id indicado.
    Cada elemento de la lista es otra lista con:
    [category, offeringID, offering, parent_offeringID, parent_offering_name, characteristics_list]

    Además, normaliza AA Amount dentro de characteristics_list (saca la notación científica).
    Solo trae OFFERING que son paquetes (no bonos).
    """
    global connection, listaPaquetes

    listaPaquetes = []
    cursor = connection.cursor()

    sql = """
    WITH offers AS (
        -- Una sola fila por OFFERINGID (si hay varios padres, se toma uno solo)
        SELECT
            MIN(cat.name)       AS category,
            po.PRD_OFF          AS offeringID,
            po.name             AS offering,
            MIN(parent.PRD_OFF) AS parent_offeringID,
            MIN(parent.name)    AS parent_offering_name
        FROM prd_app_6800.R_PIM_PRD_OFF po
        LEFT JOIN prd_app_6800.R_PIM_PCHLD_OF_RLHP rel 
               ON po.PRD_OFF = rel.CHILD
        LEFT JOIN prd_app_6800.R_PIM_OFFR_CAT cat 
               ON cat.offr_cat = rel.PARENT
        LEFT JOIN prd_app_6800.R_PIM_POFF_CDL_RLT rel_parent 
               ON po.PRD_OFF = rel_parent.CHILD
        LEFT JOIN prd_app_6800.R_PIM_PRD_OFF parent 
               ON parent.PRD_OFF = rel_parent.PARENT
        LEFT JOIN prd_app_6800.nc_cmt_ct_modifs cr 
               ON cr.ci_id = po.PRD_OFF
        WHERE cr.activity_id = :activity_id
          AND po.name LIKE 'Paquete%'          -- ⬅️ solo paquetes, desestima bonos
        GROUP BY
            po.PRD_OFF,
            po.name
    ),

    chars AS (
        -- Características únicas por oferta
        SELECT DISTINCT
            po.PRD_OFF AS offeringID,
            chv.name   AS chr_name,

            /* 
               Traemos el primer valor no nulo entre:
               - columnas de texto (DEFAULT_*)
               - columnas numéricas convertidas a texto (DEF_LIST_*)
               No hacemos TO_NUMBER para evitar ORA-01722.
            */
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
            ) AS chr_value

        FROM prd_app_6800.R_PIM_PRD_OFF po
        LEFT JOIN prd_app_6800.R_PIM_OFFR_CHR_INV chv 
               ON po.prd_off = chv.parent_ID
        LEFT JOIN prd_app_6800.nc_cmt_ct_modifs cr 
               ON cr.ci_id = po.PRD_OFF
        WHERE cr.activity_id = :activity_id
    ),

    agg_chars AS (
        -- Lista ordenada, sin valores repetidos para cada offering
        SELECT
            c.offeringID,
            RTRIM(
                XMLCAST(
                    XMLAGG(
                        XMLELEMENT(
                            e,
                            c.chr_name || ': ' || c.chr_value || ', '
                        )
                        ORDER BY c.chr_name
                    ) AS CLOB
                ),
                ', '
            ) AS characteristics_list
        FROM chars c
        GROUP BY c.offeringID
    )

    SELECT
        o.category,
        o.offeringID,
        o.offering,
        o.parent_offeringID,
        o.parent_offering_name,
        a.characteristics_list
    FROM offers o
    LEFT JOIN agg_chars a
           ON a.offeringID = o.offeringID
    ORDER BY
        o.parent_offering_name,
        o.offering
    """

    cursor.execute(sql, activity_id=activity_id)

    for row in cursor:
        fila = list(row)
        # La última columna es CHARACTERISTICS_LIST
        fila[-1] = normalizar_aa_amount(fila[-1])
        listaPaquetes.append(fila)

    cursor.close()
    return listaPaquetes
