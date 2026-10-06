@EndUserText.label: 'Sales Report - Export Excel action'
@ObjectModel.query.implementedBy: 'ABAP:ZCL_CE_SALES_REPORT_EXP'
define root custom entity zce_sales_report_exp
{
      // Entity chỉ để gắn action ExportExcel, không có dữ liệu
  key ExportId : abap.char(1);
}
