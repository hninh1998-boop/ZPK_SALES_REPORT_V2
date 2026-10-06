CLASS zcl_ce_sales_report DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    INTERFACES if_rap_query_provider .

    TYPES tt_result   TYPE STANDARD TABLE OF zce_sales_report WITH EMPTY KEY.
    TYPES tr_plant    TYPE RANGE OF werks_d.
    TYPES tr_product  TYPE RANGE OF matnr.
    TYPES tr_company  TYPE RANGE OF bukrs.
    TYPES tr_fyear    TYPE RANGE OF gjahr.
    TYPES tr_pdate    TYPE RANGE OF zce_sales_report-postingdate.
    TYPES tr_district TYPE RANGE OF char6.

    " Bỏ số 0 ở đầu mã hàng trong filter, để so với ltrim( product, '0' ) trong get_data
    " (EQ 300004390 và pattern 3* đều khớp với mã nội bộ 000000000300004390)
    CLASS-METHODS conv_product_range
      IMPORTING it_range          TYPE if_rap_query_filter=>tt_range_option
      RETURNING VALUE(rr_product) TYPE tr_product.

    " Lấy dữ liệu báo cáo theo filter (chưa aggregation / sort / paging)
    " Dùng chung cho query của custom entity và action ExportExcel
    CLASS-METHODS get_data
      IMPORTING ir_plant         TYPE tr_plant    OPTIONAL
                ir_product       TYPE tr_product  OPTIONAL
                ir_company       TYPE tr_company
                ir_fyear         TYPE tr_fyear    OPTIONAL
                ir_pdate         TYPE tr_pdate
                ir_district      TYPE tr_district OPTIONAL
      RETURNING VALUE(rt_result) TYPE tt_result.
  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_ce_sales_report IMPLEMENTATION.


  METHOD if_rap_query_provider~select.

    DATA lt_result TYPE tt_result.
    DATA lv_total  TYPE i.
    DATA lv_top    TYPE i.
    DATA lv_skip   TYPE i.

    " ── 1. LẤY FILTER ──────────────────────────────────────
    DATA(lt_filters) = io_request->get_filter( )->get_as_ranges( ).

    DATA lr_plant      TYPE tr_plant.
    DATA lr_product    TYPE tr_product.
    DATA lr_company    TYPE tr_company.
    DATA lr_fyear      TYPE tr_fyear.
    DATA lr_pdate      TYPE tr_pdate.
    DATA lr_district   TYPE tr_district.

    LOOP AT lt_filters INTO DATA(ls_filter).
      CASE ls_filter-name.
        WHEN 'PLANT'.
          lr_plant    = CORRESPONDING #( ls_filter-range ).
        WHEN 'PRODUCT'.
          lr_product  = conv_product_range( ls_filter-range ).
        WHEN 'COMPANYCODE'.
          lr_company  = CORRESPONDING #( ls_filter-range ).
        WHEN 'FISCALYEAR'.
          lr_fyear    = CORRESPONDING #( ls_filter-range ).
        WHEN 'POSTINGDATE'.
          lr_pdate    = CORRESPONDING #( ls_filter-range ).
        WHEN 'SALESDISTRICT'.
          lr_district = CORRESPONDING #( ls_filter-range ).
      ENDCASE.
    ENDLOOP.

    " Validate mandatory filters
    IF lr_company IS INITIAL OR lr_pdate IS INITIAL.
      io_response->set_total_number_of_records( 0 ).
      io_response->set_data( lt_result ).
      RETURN.
    ENDIF.

    " ── 2. - 4. LẤY DỮ LIỆU ────────────────────────────────
    lt_result = get_data( ir_plant    = lr_plant
                          ir_product  = lr_product
                          ir_company  = lr_company
                          ir_fyear    = lr_fyear
                          ir_pdate    = lr_pdate
                          ir_district = lr_district ).

    " ── 5. HANDLE AGGREGATION ───────────────────────────────
    DATA(lo_aggregation) = io_request->get_aggregation( ).
    DATA(lt_group_by)    = lo_aggregation->get_grouped_elements( ).
    DATA(lt_agg_elems)   = lo_aggregation->get_aggregated_elements( ).

    IF lt_group_by IS NOT INITIAL OR lt_agg_elems IS NOT INITIAL.

      DATA lt_aggregated TYPE HASHED TABLE OF zce_sales_report
        WITH UNIQUE KEY plant product companycode fiscalyear postingdate
                        salesdistrict companycodecurrency currencyusd productgroup.

      LOOP AT lt_result INTO DATA(ls_agg_in).

        DATA(lv_key_plant) = COND werks_d(
          WHEN line_exists( lt_group_by[ table_line = 'PLANT' ] )
          THEN ls_agg_in-plant ELSE '' ).

        DATA(lv_key_product) = COND matnr(
          WHEN line_exists( lt_group_by[ table_line = 'PRODUCT' ] )
          THEN ls_agg_in-product ELSE '' ).

        DATA(lv_key_company) = COND bukrs(
          WHEN line_exists( lt_group_by[ table_line = 'COMPANYCODE' ] )
          THEN ls_agg_in-companycode ELSE '' ).

        DATA(lv_key_fyear) = COND gjahr(
          WHEN line_exists( lt_group_by[ table_line = 'FISCALYEAR' ] )
          THEN ls_agg_in-fiscalyear ELSE '' ).

        " Không ELSE: kiểu ngày phải để initial (00000000), gán '' sẽ ra ngày không hợp lệ
        DATA(lv_key_pdate) = COND zce_sales_report-postingdate(
          WHEN line_exists( lt_group_by[ table_line = 'POSTINGDATE' ] )
          THEN ls_agg_in-postingdate ).

        DATA(lv_key_district) = COND char6(
          WHEN line_exists( lt_group_by[ table_line = 'SALESDISTRICT' ] )
          THEN ls_agg_in-salesdistrict ELSE '' ).

        DATA(lv_key_cccy) = COND waers(
          WHEN line_exists( lt_group_by[ table_line = 'COMPANYCODECURRENCY' ] )
          THEN ls_agg_in-companycodecurrency ELSE '' ).

        DATA(lv_key_usdccy) = COND waers(
          WHEN line_exists( lt_group_by[ table_line = 'CURRENCYUSD' ] )
          THEN ls_agg_in-currencyusd ELSE '' ).

        DATA(lv_key_productgroup) = COND char9(
          WHEN line_exists( lt_group_by[ table_line = 'PRODUCTGROUP' ] )
          THEN ls_agg_in-ProductGroup ELSE '' ).

        ASSIGN lt_aggregated[ plant               = lv_key_plant
                              product             = lv_key_product
                              companycode         = lv_key_company
                              fiscalyear          = lv_key_fyear
                              postingdate         = lv_key_pdate
                              salesdistrict       = lv_key_district
                              companycodecurrency = lv_key_cccy
                              currencyusd         = lv_key_usdccy
                              productgroup        = lv_key_productgroup ]
          TO FIELD-SYMBOL(<ls_agg>).

        IF sy-subrc = 0.
          <ls_agg>-quantity   += ls_agg_in-quantity.
          <ls_agg>-revenueusd += ls_agg_in-revenueusd.
          <ls_agg>-revenuevnd += ls_agg_in-revenuevnd.
          <ls_agg>-cogsvnd    += ls_agg_in-cogsvnd.
        ELSE.
          DATA(ls_agg_out)               = ls_agg_in.
          ls_agg_out-plant               = lv_key_plant.
          ls_agg_out-product             = lv_key_product.
          ls_agg_out-companycode         = lv_key_company.
          ls_agg_out-fiscalyear          = lv_key_fyear.
          ls_agg_out-postingdate         = lv_key_pdate.
          ls_agg_out-salesdistrict       = lv_key_district.
          ls_agg_out-companycodecurrency = lv_key_cccy.
          ls_agg_out-currencyusd         = lv_key_usdccy.
          ls_agg_out-ProductGroup        = lv_key_productgroup.
          INSERT ls_agg_out INTO TABLE lt_aggregated.
        ENDIF.

      ENDLOOP.

      lt_result = lt_aggregated.

    ENDIF.

    " ── 6a. RESPONSE ────────────────────────────────
    lv_total  = lines( lt_result ).

    IF io_request->is_total_numb_of_rec_requested( ).
      io_response->set_total_number_of_records( CONV int8( lv_total ) ).
    ENDIF.

    " ── 6b. HANDLE SORT ─────────────────────────────────────
    DATA(lt_sort) = io_request->get_sort_elements( ).
    IF lt_sort IS NOT INITIAL.
      DATA lt_sort_order TYPE abap_sortorder_tab.
      LOOP AT lt_sort INTO DATA(ls_sort).
        APPEND VALUE #(
            name       = ls_sort-element_name
            descending = ls_sort-descending
        ) TO lt_sort_order.
      ENDLOOP.
      SORT lt_result BY (lt_sort_order).
    ENDIF.

    " ── 6c. PAGING ──────────────────────────────────────────
    lv_skip = io_request->get_paging( )->get_offset( ).
    lv_top  = io_request->get_paging( )->get_page_size( ).

    IF lv_top = if_rap_query_paging=>page_size_unlimited.
      lv_top = lv_total.
    ENDIF.

    IF lv_skip > 0.
      DELETE lt_result TO lv_skip.
    ENDIF.

    IF lv_top < lines( lt_result ).
      DELETE lt_result FROM lv_top + 1.
    ENDIF.

    io_response->set_data( lt_result ).

  ENDMETHOD.


  METHOD conv_product_range.

    LOOP AT it_range INTO DATA(ls_range_prod).
      rr_product = VALUE #( BASE rr_product (
        sign   = ls_range_prod-sign
        option = ls_range_prod-option
        low    = shift_left( val = ls_range_prod-low  sub = '0' )
        high   = shift_left( val = ls_range_prod-high sub = '0' )
      ) ).
    ENDLOOP.

  ENDMETHOD.


  METHOD get_data.

    " ── 2. SELECT riêng cho COGSVND (sản phẩm 1xxx) ────────────
    TYPES: BEGIN OF lty_cogs,
             plant         TYPE werks_d,
             product       TYPE matnr,
             companycode   TYPE bukrs,
             fiscalyear    TYPE gjahr,
             postingdate   TYPE zce_sales_report-postingdate,
             salesdistrict TYPE char6,
             profitcenter  TYPE prctr,
             baseunit      TYPE meins,
             cogsvnd       TYPE p LENGTH 15 DECIMALS 2,
           END OF lty_cogs.

    DATA lt_cogs TYPE STANDARD TABLE OF lty_cogs WITH EMPTY KEY.

    " Lấy COGSVND cho sản phẩm 1xxx có goodsmovementtype X49/X50
    SELECT gli~plant,
           gli~product,
           gli~companycode,
           gli~fiscalyear,
           gli~postingdate,
           gli~salesdistrict,
           gli~profitcenter,
           gli~baseunit,
           SUM( gli~amountincompanycodecurrency ) AS cogsvnd
      FROM i_glaccountlineitemsemtag AS gli
      INNER JOIN i_materialdocumentitem_2 AS mat
        ON mat~materialdocument = gli~referencedocument
       AND mat~goodsmovementtype IN ( 'X49', 'X50' )
      WHERE gli~companycode        IN @ir_company
        AND gli~fiscalyear         IN @ir_fyear
        AND gli~postingdate        IN @ir_pdate
        AND gli~plant              IN @ir_plant
        AND ltrim( gli~product, '0' ) IN @ir_product
        AND gli~glaccount LIKE '632%'
        AND gli~ledger             = '0L'
        AND gli~glaccounthierarchy = 'ZPL'
        AND gli~semantictag        = 'PL_RESULT'
        AND ltrim( gli~product, '0' ) LIKE '1%'
      GROUP BY gli~plant, gli~product, gli~companycode,
               gli~fiscalyear, gli~postingdate,
               gli~salesdistrict, gli~profitcenter, gli~baseunit
      INTO TABLE @lt_cogs.

    " ── 2a. SELECT chính ───────────────────────────────────────
    SELECT
        gli~plant,
        gli~product,
        gli~companycode,
        gli~fiscalyear,
        gli~salesdistrict,
        gli~postingdate,
        gli~profitcenter,
        ptxt~productname,
        gli~soldproductgroup AS productgroup,
        ProGrpTxt~productgroupname,
        sdtxt~salesdistrictname,
        gli~companycodecurrency,
        SUM( CASE
            WHEN gli~glaccount LIKE '511%'  THEN gli~quantity
            WHEN gli~glaccount LIKE '5212%' THEN gli~quantity
            ELSE CAST( 0 AS QUAN( 13, 3 ) )
        END ) AS quantity,
        gli~BaseUnit,
        SUM( CASE
            WHEN gli~glaccount LIKE '511%'  AND gli~transactioncurrency = 'USD'
            THEN gli~amountintransactioncurrency
            WHEN gli~glaccount LIKE '5212%' AND gli~transactioncurrency = 'USD'
            THEN gli~amountintransactioncurrency
            ELSE CAST( 0 AS CURR( 15, 2 ) )
        END ) AS revenueusd,
        SUM( CASE
            WHEN gli~glaccount LIKE '511%'  THEN gli~amountincompanycodecurrency
            WHEN gli~glaccount LIKE '5212%' THEN gli~amountincompanycodecurrency
            ELSE CAST( 0 AS CURR( 15, 2 ) )
        END ) AS revenuevnd,
        SUM( CASE
            WHEN gli~glaccount LIKE '632%'
                AND gli~ledger             = '0L'
                AND gli~glaccounthierarchy = 'ZPL'
                AND gli~semantictag        = 'PL_RESULT'
                AND ( ltrim( gli~product, '0' ) NOT LIKE '1%'
                      OR gli~product IS NULL )
                AND gli~salesdocument IS NOT NULL
                AND gli~salesdocument IS NOT INITIAL
            THEN gli~amountincompanycodecurrency
            ELSE CAST( 0 AS CURR( 15, 2 ) )
        END ) AS cogsvnd

        FROM i_glaccountlineitemsemtag AS gli

        LEFT OUTER JOIN i_producttext AS ptxt
        ON  ptxt~product  = gli~product
        AND ptxt~language = @sy-langu

        LEFT OUTER JOIN i_salesdistricttext AS sdtxt
        ON  sdtxt~salesdistrict = gli~salesdistrict
        AND sdtxt~language      = @sy-langu

        LEFT OUTER JOIN i_journalentry AS revdoc
            ON  gli~companycode = revdoc~companycode
            AND gli~fiscalyear  = revdoc~fiscalyear
            AND ( gli~isreversed = 'X' OR gli~isreversal = 'X' )
            AND (
            revdoc~originalreferencedocument = gli~reversalreferencedocument
            OR revdoc~originalreferencedocument =
            concat( gli~reversalreferencedocument, gli~fiscalyear )
            OR revdoc~originalreferencedocument =
            concat( gli~reversalreferencedocument,
            concat( gli~companycode, gli~fiscalyear ) )
            )

        LEFT OUTER JOIN i_cnsldtnproductgroupvh AS ProGrpTxt
            ON ProGrpTxt~productgroup = gli~soldproductgroup

        WHERE gli~companycode        IN @ir_company
        AND gli~fiscalyear         IN @ir_fyear
        AND gli~postingdate        IN @ir_pdate
        AND gli~plant              IN @ir_plant
        AND ltrim( gli~product, '0' ) IN @ir_product
        AND gli~salesdistrict      IN @ir_district
        AND gli~ledger             =  '0L'
        AND gli~glaccounthierarchy =  'ZPL'
        AND gli~semantictag        =  'PL_RESULT'
        AND ( gli~glaccount LIKE '511%'
        OR gli~glaccount LIKE '5212%'
        OR gli~glaccount LIKE '632%' )
        AND gli~product            <> ''
        AND gli~product            IS NOT NULL
        AND (    revdoc~accountingdocument IS NULL
        OR revdoc~accountingdocument =  ''
        OR (     revdoc~accountingdocument IS NOT NULL
        AND revdoc~accountingdocument <> ''
        AND gli~fiscalperiod <> revdoc~fiscalperiod ) )

        GROUP BY
        gli~BaseUnit,
        gli~plant,
        gli~product,
        gli~companycode,
        gli~fiscalyear,
        gli~salesdistrict,
        gli~postingdate,
        gli~profitcenter,
        ptxt~productname,
        gli~soldproductgroup,
        sdtxt~salesdistrictname,
        gli~companycodecurrency,
        ProGrpTxt~productgroupname

      INTO TABLE @DATA(lt_db).

    " Merge COGSVND (1xxx) vào lt_db – mỗi dòng COGS chỉ cộng đúng 1 lần
    " Không tìm thấy → thêm dòng mới (chỉ có COGS, không có revenue)
    LOOP AT lt_cogs INTO DATA(ls_cogs).
      READ TABLE lt_db ASSIGNING FIELD-SYMBOL(<ls_merge>)
          WITH KEY plant         = ls_cogs-plant
                   product       = ls_cogs-product
                   companycode   = ls_cogs-companycode
                   fiscalyear    = ls_cogs-fiscalyear
                   postingdate   = ls_cogs-postingdate
                   salesdistrict = ls_cogs-salesdistrict
                   profitcenter  = ls_cogs-profitcenter.
      IF sy-subrc = 0.
        <ls_merge>-cogsvnd += ls_cogs-cogsvnd.
      ELSE.
        APPEND VALUE #(
            plant         = ls_cogs-plant
            product       = ls_cogs-product
            companycode   = ls_cogs-companycode
            fiscalyear    = ls_cogs-fiscalyear
            postingdate   = ls_cogs-postingdate
            salesdistrict = ls_cogs-salesdistrict
            profitcenter  = ls_cogs-profitcenter
            baseunit      = ls_cogs-baseunit
            cogsvnd       = ls_cogs-cogsvnd
        ) TO lt_db.
      ENDIF.
    ENDLOOP.

    " ── 2b. POST-PROCESS: Fix Plant nếu gli~plant trống ─────
    DATA lt_prctr_range TYPE RANGE OF prctr.
    DATA lt_prod_range  TYPE RANGE OF matnr.

    LOOP AT lt_db INTO DATA(ls_chk)
      WHERE plant IS INITIAL OR plant = ''.
      APPEND VALUE #( sign = 'I' option = 'EQ' low = ls_chk-profitcenter ) TO lt_prctr_range.
      APPEND VALUE #( sign = 'I' option = 'EQ' low = ls_chk-product )      TO lt_prod_range.
    ENDLOOP.

    SORT lt_prctr_range BY low.
    DELETE ADJACENT DUPLICATES FROM lt_prctr_range COMPARING low.
    SORT lt_prod_range BY low.
    DELETE ADJACENT DUPLICATES FROM lt_prod_range COMPARING low.

    TYPES: BEGIN OF lty_pp,
             product      TYPE matnr,
             profitcenter TYPE prctr,
             plant        TYPE werks_d,
           END OF lty_pp,
           BEGIN OF lty_map,
             profitcenter TYPE prctr,
             plant        TYPE werks_d,
           END OF lty_map.
    DATA lt_pp  TYPE SORTED TABLE OF lty_pp  WITH NON-UNIQUE KEY product profitcenter.
    DATA lt_map TYPE SORTED TABLE OF lty_map WITH NON-UNIQUE KEY profitcenter.

    IF lt_prctr_range IS NOT INITIAL.
      SELECT product, profitcenter, plant
        FROM i_productplantbasic
        WHERE product      IN @lt_prod_range
          AND profitcenter IN @lt_prctr_range
        INTO TABLE @lt_pp.

      SELECT profitcenter, plant
        FROM zi_plt_cc_prctr
        WHERE profitcenter IN @lt_prctr_range
        INTO TABLE @lt_map.
    ENDIF.

    LOOP AT lt_db ASSIGNING FIELD-SYMBOL(<ls_db>)
      WHERE plant IS INITIAL OR plant = ''.

      " Case 1: lấy từ pp (plant đầu tiên match product + profitcenter)
      READ TABLE lt_pp INTO DATA(ls_pp)
        WITH KEY product      = <ls_db>-product
                 profitcenter = <ls_db>-profitcenter.
      IF sy-subrc = 0 AND ls_pp-plant IS NOT INITIAL AND ls_pp-plant <> ''.
        <ls_db>-plant = ls_pp-plant.
        CONTINUE.
      ENDIF.

      " Case 2: fallback từ map
      READ TABLE lt_map INTO DATA(ls_map)
        WITH KEY profitcenter = <ls_db>-profitcenter.
      IF sy-subrc = 0 AND ls_map-plant IS NOT INITIAL AND ls_map-plant <> ''.
        <ls_db>-plant = ls_map-plant.
      ENDIF.
    ENDLOOP.

    " ── 2c. POST-PROCESS: Fix ProductGroup + ProductGroupName ──
    DATA lt_prod_range2 TYPE RANGE OF matnr.
    LOOP AT lt_db INTO DATA(ls_chk2).
      APPEND VALUE #( sign = 'I' option = 'EQ' low = ls_chk2-product )
          TO lt_prod_range2.
    ENDLOOP.
    SORT lt_prod_range2 BY low.
    DELETE ADJACENT DUPLICATES FROM lt_prod_range2 COMPARING low.

    TYPES: BEGIN OF lty_prod_grp,
             product      TYPE matnr,
             productgroup TYPE matkl,
           END OF lty_prod_grp.
    DATA lt_prod_grp TYPE SORTED TABLE OF lty_prod_grp WITH NON-UNIQUE KEY product.

    IF lt_prod_range2 IS NOT INITIAL.
      SELECT product, productgroup
          FROM i_product
          WHERE product IN @lt_prod_range2
          INTO TABLE @lt_prod_grp.
    ENDIF.

    DATA lt_pg_range TYPE RANGE OF matkl.
    LOOP AT lt_db INTO DATA(ls_chk3).
      IF ls_chk3-productgroup IS NOT INITIAL AND ls_chk3-productgroup <> ''.
        APPEND VALUE #( sign = 'I' option = 'EQ' low = ls_chk3-productgroup )
            TO lt_pg_range.
      ENDIF.
    ENDLOOP.
    LOOP AT lt_prod_grp INTO DATA(ls_pg_tmp).
      APPEND VALUE #( sign = 'I' option = 'EQ' low = ls_pg_tmp-productgroup )
          TO lt_pg_range.
    ENDLOOP.
    SORT lt_pg_range BY low.
    DELETE ADJACENT DUPLICATES FROM lt_pg_range COMPARING low.

    TYPES: BEGIN OF lty_pg_text,
             productgroup     TYPE matkl,
             productgroupname TYPE char20,
           END OF lty_pg_text.
    DATA lt_pg_text TYPE SORTED TABLE OF lty_pg_text WITH NON-UNIQUE KEY productgroup.

    IF lt_pg_range IS NOT INITIAL.
      SELECT productgroup, productgroupname
          FROM i_cnsldtnproductgroupvh
          WHERE productgroup IN @lt_pg_range
          INTO TABLE @lt_pg_text.
    ENDIF.

    LOOP AT lt_db ASSIGNING FIELD-SYMBOL(<ls_db2>).
      " Fix productgroup nếu trống → lấy từ i_product
      IF <ls_db2>-productgroup IS INITIAL OR <ls_db2>-productgroup = ''.
        READ TABLE lt_prod_grp INTO DATA(ls_pg)
            WITH KEY product = <ls_db2>-product.
        IF sy-subrc = 0.
          <ls_db2>-productgroup = ls_pg-productgroup.
        ENDIF.
      ENDIF.

      " Nếu productgroupname vẫn trống → lookup lại theo productgroup mới fill
      IF <ls_db2>-productgroupname IS INITIAL OR <ls_db2>-productgroupname = ''.
        READ TABLE lt_pg_text INTO DATA(ls_pg_name)
            WITH KEY productgroup = <ls_db2>-productgroup.
        IF sy-subrc = 0.
          <ls_db2>-productgroupname = ls_pg_name-productgroupname.
        ENDIF.
      ENDIF.
    ENDLOOP.

    " ── 3. MAP TO RESULT ────────────────────────────────────
    LOOP AT lt_db INTO DATA(ls_db).
      APPEND VALUE zce_sales_report(
        plant               = ls_db-plant
        product             = ls_db-product
        companycode         = ls_db-companycode
        fiscalyear          = ls_db-fiscalyear
        salesdistrict       = ls_db-salesdistrict
        postingdate         = ls_db-postingdate
        profitcenter        = ls_db-profitcenter
        productname         = ls_db-productname
        productgroup        = ls_db-productgroup
        productgroupname    = ls_db-productgroupname
        salesdistrictname   = ls_db-salesdistrictname
        quantity            = - ls_db-quantity
        revenueusd          = - ls_db-revenueusd
        currencyusd         = 'USD'
        revenuevnd          = - ls_db-revenuevnd
        cogsvnd             = ls_db-cogsvnd
        companycodecurrency = ls_db-companycodecurrency
        BaseUnit            = ls_db-BaseUnit
      ) TO rt_result.
    ENDLOOP.

    " ── 4. FILTER ZERO ROWS ─────────────────────────────────
    DELETE rt_result WHERE quantity   = 0
                       AND revenueusd = 0
                       AND revenuevnd = 0
                       AND cogsvnd    = 0.

  ENDMETHOD.
ENDCLASS.


