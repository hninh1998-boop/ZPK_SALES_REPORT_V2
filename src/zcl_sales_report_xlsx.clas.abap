CLASS zcl_sales_report_xlsx DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    " Sinh file xlsx "BÁO CÁO BÁN HÀNG" có format (nền vàng, viền, merge, number format)
    " xlsx = file zip chứa các file XML → tự viết XML rồi nén bằng cl_abap_zip
    CLASS-METHODS build
      IMPORTING it_data        TYPE zcl_ce_sales_report=>tt_result
                iv_subtitle    TYPE string
      RETURNING VALUE(rv_xlsx) TYPE xstring.

  PROTECTED SECTION.
  PRIVATE SECTION.

    TYPES ty_num TYPE p LENGTH 16 DECIMALS 3.

    " 1 dòng trên file = 1 tổ hợp của các cột dimension hiển thị
    TYPES: BEGIN OF ty_row,
             plant             TYPE zce_sales_report-plant,
             product           TYPE zce_sales_report-product,
             productgroup      TYPE zce_sales_report-productgroup,
             salesdistrict     TYPE zce_sales_report-salesdistrict,
             productname       TYPE zce_sales_report-productname,
             productgroupname  TYPE zce_sales_report-productgroupname,
             salesdistrictname TYPE zce_sales_report-salesdistrictname,
             quantity          TYPE ty_num,
             revenueusd        TYPE ty_num,
             revenuevnd        TYPE ty_num,
             cogsvnd           TYPE ty_num,
           END OF ty_row.
    TYPES tt_row TYPE SORTED TABLE OF ty_row
                 WITH UNIQUE KEY plant product productgroup salesdistrict.

    TYPES: BEGIN OF ty_factor,
             currency TYPE zce_sales_report-companycodecurrency,
             factor   TYPE decfloat34,
           END OF ty_factor.
    TYPES tt_factor TYPE HASHED TABLE OF ty_factor WITH UNIQUE KEY currency.

    " Index trong <cellXfs> của styles.xml
    CONSTANTS: BEGIN OF gc_style,
                 title      TYPE i VALUE 1,
                 subtitle   TYPE i VALUE 2,
                 header     TYPE i VALUE 3,
                 text       TYPE i VALUE 4,
                 quantity   TYPE i VALUE 5,
                 amount_usd TYPE i VALUE 6,
                 amount_vnd TYPE i VALUE 7,
                 total_text TYPE i VALUE 8,
                 total_qty  TYPE i VALUE 9,
                 total_usd  TYPE i VALUE 10,
                 total_vnd  TYPE i VALUE 11,
               END OF gc_style.

    CONSTANTS gc_header_row TYPE i VALUE 4.

    CLASS-METHODS aggregate
      IMPORTING it_data        TYPE zcl_ce_sales_report=>tt_result
      RETURNING VALUE(rt_rows) TYPE tt_row.

    " Amount kiểu CURR lưu theo 2 số lẻ → nhân hệ số để ra giá trị thật (VND: x100)
    CLASS-METHODS get_currency_factor
      IMPORTING iv_currency      TYPE zce_sales_report-companycodecurrency
      CHANGING  ct_factor        TYPE tt_factor
      RETURNING VALUE(rv_factor) TYPE decfloat34.

    CLASS-METHODS sheet_xml
      IMPORTING it_rows       TYPE tt_row
                iv_subtitle   TYPE string
      RETURNING VALUE(rv_xml) TYPE string.

    CLASS-METHODS styles_xml
      RETURNING VALUE(rv_xml) TYPE string.

    CLASS-METHODS text_cell
      IMPORTING iv_col        TYPE string
                iv_row        TYPE i
                iv_style      TYPE i
                iv_value      TYPE csequence OPTIONAL
      RETURNING VALUE(rv_xml) TYPE string.

    CLASS-METHODS number_cell
      IMPORTING iv_col        TYPE string
                iv_row        TYPE i
                iv_style      TYPE i
                iv_value      TYPE ty_num
      RETURNING VALUE(rv_xml) TYPE string.

    CLASS-METHODS code_with_name
      IMPORTING iv_code        TYPE csequence
                iv_name        TYPE csequence
      RETURNING VALUE(rv_text) TYPE string.

    CLASS-METHODS to_utf8
      IMPORTING iv_string         TYPE string
      RETURNING VALUE(rv_xstring) TYPE xstring.

ENDCLASS.



CLASS zcl_sales_report_xlsx IMPLEMENTATION.


  METHOD build.

    DATA(lt_rows) = aggregate( it_data ).

    DATA(lv_xml_decl) = |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>|.

    DATA(lv_content_types) = lv_xml_decl
      && |<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">|
      && |<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>|
      && |<Default Extension="xml" ContentType="application/xml"/>|
      && |<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>|
      && |<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>|
      && |<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>|
      && |</Types>|.

    DATA(lv_rels) = lv_xml_decl
      && |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">|
      && |<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>|
      && |</Relationships>|.

    DATA(lv_workbook) = lv_xml_decl
      && |<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"|
      && | xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">|
      && |<sheets><sheet name="Báo cáo bán hàng" sheetId="1" r:id="rId1"/></sheets>|
      && |</workbook>|.

    DATA(lv_workbook_rels) = lv_xml_decl
      && |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">|
      && |<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>|
      && |<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>|
      && |</Relationships>|.

    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->add( name = '[Content_Types].xml'        content = to_utf8( lv_content_types ) ).
    lo_zip->add( name = '_rels/.rels'                content = to_utf8( lv_rels ) ).
    lo_zip->add( name = 'xl/workbook.xml'            content = to_utf8( lv_workbook ) ).
    lo_zip->add( name = 'xl/_rels/workbook.xml.rels' content = to_utf8( lv_workbook_rels ) ).
    lo_zip->add( name = 'xl/styles.xml'              content = to_utf8( lv_xml_decl && styles_xml( ) ) ).
    lo_zip->add( name = 'xl/worksheets/sheet1.xml'
                 content = to_utf8( lv_xml_decl && sheet_xml( it_rows     = lt_rows
                                                              iv_subtitle = iv_subtitle ) ) ).

    rv_xlsx = lo_zip->save( ).

  ENDMETHOD.


  METHOD aggregate.

    DATA lt_factor TYPE tt_factor.

    LOOP AT it_data INTO DATA(ls_data).

      ASSIGN rt_rows[ plant         = ls_data-plant
                      product       = ls_data-product
                      productgroup  = ls_data-productgroup
                      salesdistrict = ls_data-salesdistrict ]
        TO FIELD-SYMBOL(<ls_row>).

      IF sy-subrc <> 0.
        INSERT VALUE #( plant             = ls_data-plant
                        product           = ls_data-product
                        productgroup      = ls_data-productgroup
                        salesdistrict     = ls_data-salesdistrict
                        productname       = ls_data-productname
                        productgroupname  = ls_data-productgroupname
                        salesdistrictname = ls_data-salesdistrictname )
          INTO TABLE rt_rows ASSIGNING <ls_row>.
      ENDIF.

      DATA(lv_factor_usd) = get_currency_factor( EXPORTING iv_currency = ls_data-currencyusd
                                                 CHANGING  ct_factor   = lt_factor ).
      DATA(lv_factor_cc)  = get_currency_factor( EXPORTING iv_currency = ls_data-companycodecurrency
                                                 CHANGING  ct_factor   = lt_factor ).

      <ls_row>-quantity   += ls_data-quantity.
      <ls_row>-revenueusd += ls_data-revenueusd * lv_factor_usd.
      <ls_row>-revenuevnd += ls_data-revenuevnd * lv_factor_cc.
      <ls_row>-cogsvnd    += ls_data-cogsvnd    * lv_factor_cc.

    ENDLOOP.

  ENDMETHOD.


  METHOD get_currency_factor.

    READ TABLE ct_factor INTO DATA(ls_factor) WITH TABLE KEY currency = iv_currency.
    IF sy-subrc = 0.
      rv_factor = ls_factor-factor.
      RETURN.
    ENDIF.

    rv_factor = 1.

    IF iv_currency IS NOT INITIAL.
      SELECT SINGLE decimals
        FROM i_currency
        WHERE currency = @iv_currency
        INTO @DATA(lv_decimals).
      IF sy-subrc = 0.
        rv_factor = CONV decfloat34( 10 ) ** ( 2 - lv_decimals ).
      ENDIF.
    ENDIF.

    INSERT VALUE #( currency = iv_currency factor = rv_factor ) INTO TABLE ct_factor.

  ENDMETHOD.


  METHOD sheet_xml.

    DATA ls_total TYPE ty_row.
    DATA lv_rows  TYPE string.

    " ── Row 1: Tiêu đề / Row 2: Tháng - Năm (Row 3 để trống) ──
    lv_rows = |<row r="1" ht="20" customHeight="1">|
           && text_cell( iv_col = `A` iv_row = 1 iv_style = gc_style-title iv_value = `BÁO CÁO BÁN HÀNG` )
           && |</row>|
           && |<row r="2">|
           && text_cell( iv_col = `A` iv_row = 2 iv_style = gc_style-subtitle iv_value = iv_subtitle )
           && |</row>|.

    " ── Row 4: Header ───────────────────────────────────────
    DATA(lv_row) = gc_header_row.
    lv_rows = lv_rows
           && |<row r="{ lv_row }" ht="30" customHeight="1">|
           && text_cell( iv_col = `A` iv_row = lv_row iv_style = gc_style-header iv_value = `STT` )
           && text_cell( iv_col = `B` iv_row = lv_row iv_style = gc_style-header iv_value = `Tên nhà máy` )
           && text_cell( iv_col = `C` iv_row = lv_row iv_style = gc_style-header iv_value = `Mã hàng` )
           && text_cell( iv_col = `D` iv_row = lv_row iv_style = gc_style-header iv_value = `Tên hàng` )
           && text_cell( iv_col = `E` iv_row = lv_row iv_style = gc_style-header iv_value = `Loại hàng` )
           && text_cell( iv_col = `F` iv_row = lv_row iv_style = gc_style-header iv_value = `Số lượng` )
           && text_cell( iv_col = `G` iv_row = lv_row iv_style = gc_style-header iv_value = `Doanh thu USD` )
           && text_cell( iv_col = `H` iv_row = lv_row iv_style = gc_style-header iv_value = `Doanh thu VND` )
           && text_cell( iv_col = `I` iv_row = lv_row iv_style = gc_style-header iv_value = `Giá vốn VND` )
           && text_cell( iv_col = `J` iv_row = lv_row iv_style = gc_style-header iv_value = `Thị trường xuất khẩu` )
           && |</row>|.

    " ── Data rows ───────────────────────────────────────────
    LOOP AT it_rows INTO DATA(ls_row).
      DATA(lv_stt) = sy-tabix.
      lv_row += 1.

      DATA(lv_product) = CONV string( ls_row-product ).
      SHIFT lv_product LEFT DELETING LEADING '0'.

      lv_rows = lv_rows
             && |<row r="{ lv_row }">|
             && number_cell( iv_col = `A` iv_row = lv_row iv_style = gc_style-text       iv_value = CONV #( lv_stt ) )
             && text_cell(   iv_col = `B` iv_row = lv_row iv_style = gc_style-text       iv_value = ls_row-plant )
             && text_cell(   iv_col = `C` iv_row = lv_row iv_style = gc_style-text       iv_value = lv_product )
             && text_cell(   iv_col = `D` iv_row = lv_row iv_style = gc_style-text       iv_value = ls_row-productname )
             && text_cell(   iv_col = `E` iv_row = lv_row iv_style = gc_style-text
                             iv_value = code_with_name( iv_code = ls_row-productgroup
                                                        iv_name = ls_row-productgroupname ) )
             && number_cell( iv_col = `F` iv_row = lv_row iv_style = gc_style-quantity   iv_value = ls_row-quantity )
             && number_cell( iv_col = `G` iv_row = lv_row iv_style = gc_style-amount_usd iv_value = ls_row-revenueusd )
             && number_cell( iv_col = `H` iv_row = lv_row iv_style = gc_style-amount_vnd iv_value = ls_row-revenuevnd )
             && number_cell( iv_col = `I` iv_row = lv_row iv_style = gc_style-amount_vnd iv_value = ls_row-cogsvnd )
             && text_cell(   iv_col = `J` iv_row = lv_row iv_style = gc_style-text
                             iv_value = code_with_name( iv_code = ls_row-salesdistrict
                                                        iv_name = ls_row-salesdistrictname ) )
             && |</row>|.

      ls_total-quantity   += ls_row-quantity.
      ls_total-revenueusd += ls_row-revenueusd.
      ls_total-revenuevnd += ls_row-revenuevnd.
      ls_total-cogsvnd    += ls_row-cogsvnd.
    ENDLOOP.

    " ── Total row ───────────────────────────────────────────
    lv_row += 1.
    lv_rows = lv_rows
           && |<row r="{ lv_row }">|
           && text_cell(   iv_col = `A` iv_row = lv_row iv_style = gc_style-total_text )
           && text_cell(   iv_col = `B` iv_row = lv_row iv_style = gc_style-total_text )
           && text_cell(   iv_col = `C` iv_row = lv_row iv_style = gc_style-total_text )
           && text_cell(   iv_col = `D` iv_row = lv_row iv_style = gc_style-total_text )
           && text_cell(   iv_col = `E` iv_row = lv_row iv_style = gc_style-total_text )
           && number_cell( iv_col = `F` iv_row = lv_row iv_style = gc_style-total_qty iv_value = ls_total-quantity )
           && number_cell( iv_col = `G` iv_row = lv_row iv_style = gc_style-total_usd iv_value = ls_total-revenueusd )
           && number_cell( iv_col = `H` iv_row = lv_row iv_style = gc_style-total_vnd iv_value = ls_total-revenuevnd )
           && number_cell( iv_col = `I` iv_row = lv_row iv_style = gc_style-total_vnd iv_value = ls_total-cogsvnd )
           && text_cell(   iv_col = `J` iv_row = lv_row iv_style = gc_style-total_text )
           && |</row>|.

    rv_xml = |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">|
          " ── Column widths ──
          && |<cols>|
          && |<col min="1" max="1" width="6" customWidth="1"/>|
          && |<col min="2" max="2" width="15" customWidth="1"/>|
          && |<col min="3" max="3" width="20" customWidth="1"/>|
          && |<col min="4" max="4" width="30" customWidth="1"/>|
          && |<col min="5" max="5" width="12" customWidth="1"/>|
          && |<col min="6" max="9" width="16" customWidth="1"/>|
          && |<col min="10" max="10" width="22" customWidth="1"/>|
          && |</cols>|
          && |<sheetData>{ lv_rows }</sheetData>|
          && |<mergeCells count="2"><mergeCell ref="A1:J1"/><mergeCell ref="A2:J2"/></mergeCells>|
          && |</worksheet>|.

  ENDMETHOD.


  METHOD styles_xml.

    " Muốn đổi màu / font / viền thì sửa ở đây. Thứ tự <xf> trong <cellXfs> phải khớp gc_style
    DATA(lv_border)   = |borderId="1" applyBorder="1"|.
    DATA(lv_yellow)   = |fontId="1" fillId="2" applyFont="1" applyFill="1" { lv_border }|.
    DATA(lv_centered) = |<alignment horizontal="center" vertical="center"|.

    rv_xml = |<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">|

          && |<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.000"/></numFmts>|

          " fontId: 0 = thường, 1 = đậm, 2 = đậm cỡ 13 (tiêu đề)
          && |<fonts count="3">|
          && |<font><sz val="11"/><name val="Calibri"/></font>|
          && |<font><b/><sz val="11"/><name val="Calibri"/></font>|
          && |<font><b/><sz val="13"/><name val="Calibri"/></font>|
          && |</fonts>|

          " fillId: 0, 1 = bắt buộc của Excel, 2 = nền vàng
          && |<fills count="3">|
          && |<fill><patternFill patternType="none"/></fill>|
          && |<fill><patternFill patternType="gray125"/></fill>|
          && |<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>|
          && |</fills>|

          " borderId: 0 = không viền, 1 = viền mỏng 4 cạnh
          && |<borders count="2">|
          && |<border><left/><right/><top/><bottom/><diagonal/></border>|
          && |<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right>|
          && |<top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>|
          && |</borders>|

          && |<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>|

          " numFmtId: 3 = #,##0   4 = #,##0.00   164 = #,##0.000
          && |<cellXfs count="12">|
          " 0: mặc định
          && |<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>|
          " 1: title
          && |<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">{ lv_centered }/></xf>|
          " 2: subtitle
          && |<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">{ lv_centered }/></xf>|
          " 3: header
          && |<xf numFmtId="0" xfId="0" { lv_yellow } applyAlignment="1">{ lv_centered } wrapText="1"/></xf>|
          " 4: text
          && |<xf numFmtId="0" fontId="0" fillId="0" xfId="0" { lv_border }/>|
          " 5: quantity
          && |<xf numFmtId="164" fontId="0" fillId="0" xfId="0" { lv_border } applyNumberFormat="1"/>|
          " 6: amount_usd
          && |<xf numFmtId="4" fontId="0" fillId="0" xfId="0" { lv_border } applyNumberFormat="1"/>|
          " 7: amount_vnd
          && |<xf numFmtId="3" fontId="0" fillId="0" xfId="0" { lv_border } applyNumberFormat="1"/>|
          " 8: total_text
          && |<xf numFmtId="0" xfId="0" { lv_yellow }/>|
          " 9: total_qty
          && |<xf numFmtId="164" xfId="0" { lv_yellow } applyNumberFormat="1"/>|
          " 10: total_usd
          && |<xf numFmtId="4" xfId="0" { lv_yellow } applyNumberFormat="1"/>|
          " 11: total_vnd
          && |<xf numFmtId="3" xfId="0" { lv_yellow } applyNumberFormat="1"/>|
          && |</cellXfs>|

          && |<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>|
          && |</styleSheet>|.

  ENDMETHOD.


  METHOD text_cell.

    IF iv_value IS INITIAL.
      rv_xml = |<c r="{ iv_col }{ iv_row }" s="{ iv_style }"/>|.
      RETURN.
    ENDIF.

    DATA(lv_text) = escape( val    = CONV string( iv_value )
                            format = cl_abap_format=>e_xml_text ).

    rv_xml = |<c r="{ iv_col }{ iv_row }" s="{ iv_style }" t="inlineStr"><is><t xml:space="preserve">{ lv_text }</t></is></c>|.

  ENDMETHOD.


  METHOD number_cell.

    rv_xml = |<c r="{ iv_col }{ iv_row }" s="{ iv_style }"><v>{ iv_value }</v></c>|.

  ENDMETHOD.


  METHOD code_with_name.

    IF iv_code IS INITIAL.
      RETURN.
    ENDIF.

    rv_text = COND #( WHEN iv_name IS INITIAL
                      THEN |{ iv_code }|
                      ELSE |{ iv_code } ({ iv_name })| ).

  ENDMETHOD.


  METHOD to_utf8.

    rv_xstring = cl_abap_conv_codepage=>create_out( )->convert( iv_string ).

  ENDMETHOD.
ENDCLASS.

