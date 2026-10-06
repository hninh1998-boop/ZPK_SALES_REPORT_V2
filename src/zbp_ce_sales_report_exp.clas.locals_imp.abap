CLASS lhc_salesreportexp DEFINITION INHERITING FROM cl_abap_behavior_handler.
  PRIVATE SECTION.

    " Cấu trúc của json_string, FE gửi filter đang chọn trên màn hình, ví dụ:
    " { "CompanyCode":  [ { "sign": "I", "option": "EQ", "low": "6710" } ],
    "   "FiscalYear":   [ { "low": "2025" } ],
    "   "PostingDate":  [ { "option": "BT", "low": "20251201", "high": "20251231" } ],
    "   "Plant": [], "Product": [], "SalesDistrict": [] }
    " sign / option bỏ trống → mặc định I / EQ (BT nếu có high)
    TYPES: BEGIN OF ty_filter,
             companycode   TYPE if_rap_query_filter=>tt_range_option,
             fiscalyear    TYPE if_rap_query_filter=>tt_range_option,
             postingdate   TYPE if_rap_query_filter=>tt_range_option,
             plant         TYPE if_rap_query_filter=>tt_range_option,
             product       TYPE if_rap_query_filter=>tt_range_option,
             salesdistrict TYPE if_rap_query_filter=>tt_range_option,
           END OF ty_filter.

    METHODS get_instance_authorizations FOR INSTANCE AUTHORIZATION
      IMPORTING keys REQUEST requested_authorizations FOR salesreportexp RESULT result.

    METHODS read FOR READ
      IMPORTING keys FOR READ salesreportexp RESULT result.

    METHODS lock FOR LOCK
      IMPORTING keys FOR LOCK salesreportexp.

    METHODS exportexcel FOR MODIFY
      IMPORTING keys FOR ACTION salesreportexp~exportexcel RESULT result.

    METHODS normalize_range
      CHANGING ct_range TYPE if_rap_query_filter=>tt_range_option.

    " Dòng thứ 2 của file excel, theo khoảng Posting Date đang lọc
    METHODS get_subtitle
      IMPORTING ir_pdate           TYPE zcl_ce_sales_report=>tr_pdate
      RETURNING VALUE(rv_subtitle) TYPE string.

    " YYYYMMDD → DD/MM/YYYY
    METHODS format_date
      IMPORTING iv_date        TYPE d
      RETURNING VALUE(rv_text) TYPE string.

ENDCLASS.

CLASS lhc_salesreportexp IMPLEMENTATION.

  METHOD get_instance_authorizations.
  ENDMETHOD.

  METHOD read.
  ENDMETHOD.

  METHOD lock.
  ENDMETHOD.

  METHOD exportexcel.

    DATA ls_filter TYPE ty_filter.

    READ TABLE keys INDEX 1 INTO DATA(ls_key).
    IF sy-subrc <> 0.
      RETURN.
    ENDIF.

    " ── 1. LẤY FILTER TỪ json_string ────────────────────────
    /ui2/cl_json=>deserialize(
      EXPORTING
        json = ls_key-%param-json_string
      CHANGING
        data = ls_filter
    ).

    normalize_range( CHANGING ct_range = ls_filter-companycode ).
    normalize_range( CHANGING ct_range = ls_filter-fiscalyear ).
    normalize_range( CHANGING ct_range = ls_filter-postingdate ).
    normalize_range( CHANGING ct_range = ls_filter-plant ).
    normalize_range( CHANGING ct_range = ls_filter-product ).
    normalize_range( CHANGING ct_range = ls_filter-salesdistrict ).

    DATA(lr_company)  = CORRESPONDING zcl_ce_sales_report=>tr_company( ls_filter-companycode ).
    DATA(lr_fyear)    = CORRESPONDING zcl_ce_sales_report=>tr_fyear( ls_filter-fiscalyear ).
    DATA(lr_pdate)    = CORRESPONDING zcl_ce_sales_report=>tr_pdate( ls_filter-postingdate ).
    DATA(lr_plant)    = CORRESPONDING zcl_ce_sales_report=>tr_plant( ls_filter-plant ).
    DATA(lr_district) = CORRESPONDING zcl_ce_sales_report=>tr_district( ls_filter-salesdistrict ).
    DATA(lr_product)  = zcl_ce_sales_report=>conv_product_range( ls_filter-product ).

    " Validate mandatory filters (giống query của custom entity)
    IF lr_company IS INITIAL OR lr_fyear IS INITIAL OR lr_pdate IS INITIAL.
      APPEND VALUE #( %cid = ls_key-%cid ) TO failed-salesreportexp.
      APPEND VALUE #( %cid = ls_key-%cid
                      %msg = new_message_with_text(
                        severity = if_abap_behv_message=>severity-error
                        text     = 'Vui lòng nhập Company Code, Fiscal Year và Posting Date' )
                    ) TO reported-salesreportexp.
      RETURN.
    ENDIF.

    " ── 2. LẤY DỮ LIỆU (dùng chung logic với màn hình) ──────
    DATA(lt_data) = zcl_ce_sales_report=>get_data(
                      ir_plant    = lr_plant
                      ir_product  = lr_product
                      ir_company  = lr_company
                      ir_fyear    = lr_fyear
                      ir_pdate    = lr_pdate
                      ir_district = lr_district ).

    IF lt_data IS INITIAL.
      APPEND VALUE #( %cid = ls_key-%cid ) TO failed-salesreportexp.
      APPEND VALUE #( %cid = ls_key-%cid
                      %msg = new_message_with_text(
                        severity = if_abap_behv_message=>severity-error
                        text     = 'Không có dữ liệu để export' )
                    ) TO reported-salesreportexp.
      RETURN.
    ENDIF.

    " ── 3. DÒNG "Từ ngày ... đến ngày ..." ──────────────────
    DATA(lv_subtitle) = get_subtitle( lr_pdate ).

    " ── 4. SINH FILE XLSX ───────────────────────────────────
    DATA(lv_file_content) = zcl_sales_report_xlsx=>build(
                              it_data     = lt_data
                              iv_subtitle = lv_subtitle ).

    result = VALUE #( FOR key IN keys (
                      %cid   = key-%cid
                      %param = VALUE #( filecontent   = lv_file_content
                                        filename      = |BaoCaoBanHang_{ cl_abap_context_info=>get_system_date( ) }|
                                        fileextension = 'xlsx'
                                        mimetype      = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' )
                      ) ).

  ENDMETHOD.

  METHOD normalize_range.

    DELETE ct_range WHERE low IS INITIAL AND high IS INITIAL.

    LOOP AT ct_range ASSIGNING FIELD-SYMBOL(<ls_range>).
      IF <ls_range>-sign IS INITIAL.
        <ls_range>-sign = 'I'.
      ENDIF.
      IF <ls_range>-option IS INITIAL.
        <ls_range>-option = COND #( WHEN <ls_range>-high IS INITIAL THEN 'EQ' ELSE 'BT' ).
      ENDIF.
    ENDLOOP.

  ENDMETHOD.

  METHOD get_subtitle.

    DATA(ls_pdate) = ir_pdate[ 1 ].
    DATA(lv_low)   = format_date( ls_pdate-low ).
    DATA(lv_high)  = format_date( ls_pdate-high ).

    rv_subtitle = SWITCH #( ls_pdate-option
      WHEN 'BT'         THEN |Từ ngày { lv_low } đến ngày { lv_high }|
      WHEN 'GE' OR 'GT' THEN |Từ ngày { lv_low }|
      WHEN 'LE' OR 'LT' THEN |Đến ngày { lv_low }|
      ELSE                   |Từ ngày { lv_low } đến ngày { lv_low }| ).

  ENDMETHOD.

  METHOD format_date.

    rv_text = |{ iv_date+6(2) }/{ iv_date+4(2) }/{ iv_date(4) }|.

  ENDMETHOD.

ENDCLASS.

CLASS lsc_zce_sales_report_exp DEFINITION INHERITING FROM cl_abap_behavior_saver.
  PROTECTED SECTION.

    METHODS finalize REDEFINITION.

    METHODS check_before_save REDEFINITION.

    METHODS save REDEFINITION.

    METHODS cleanup REDEFINITION.

    METHODS cleanup_finalize REDEFINITION.

ENDCLASS.

CLASS lsc_zce_sales_report_exp IMPLEMENTATION.

  METHOD finalize.
  ENDMETHOD.

  METHOD check_before_save.
  ENDMETHOD.

  METHOD save.
  ENDMETHOD.

  METHOD cleanup.
  ENDMETHOD.

  METHOD cleanup_finalize.
  ENDMETHOD.

ENDCLASS.

