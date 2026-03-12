WITH offers AS (
    -- Una sola fila por combinación padre + oferta
    SELECT DISTINCT
        cat.name          AS category,
        po.PRD_OFF        AS offeringID,
        po.name           AS offering,
        parent.PRD_OFF    AS parent_offeringID,
        parent.name       AS parent_offering_name
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
    WHERE cr.activity_id = '9174576622413967267'
),

chars AS (
    -- Características únicas por oferta
    SELECT DISTINCT
        po.PRD_OFF         AS offeringID,
        chv.name           AS chr_name,
        chv.DEFAULT_VALUE_TEXT AS chr_value
    FROM prd_app_6800.R_PIM_PRD_OFF po
    LEFT JOIN prd_app_6800.R_PIM_OFFR_CHR_INV chv 
           ON po.prd_off = chv.parent_ID
    LEFT JOIN prd_app_6800.nc_cmt_ct_modifs cr 
           ON cr.ci_id = po.PRD_OFF
    WHERE cr.activity_id = '9174576622413967267'
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
    o.offering;
