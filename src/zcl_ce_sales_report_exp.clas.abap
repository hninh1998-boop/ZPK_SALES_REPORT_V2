CLASS zcl_ce_sales_report_exp DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    INTERFACES if_rap_query_provider .
  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_ce_sales_report_exp IMPLEMENTATION.


  METHOD if_rap_query_provider~select.

    " zce_sales_report_exp chỉ để gắn action ExportExcel → luôn trả về rỗng
    DATA lt_result TYPE STANDARD TABLE OF zce_sales_report_exp WITH EMPTY KEY.

    IF io_request->is_total_numb_of_rec_requested( ).
      io_response->set_total_number_of_records( 0 ).
    ENDIF.

    IF io_request->is_data_requested( ).
      io_response->set_data( lt_result ).
    ENDIF.

  ENDMETHOD.
ENDCLASS.

