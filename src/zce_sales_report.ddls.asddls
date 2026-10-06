@EndUserText.label: 'Custom Entity - Sales Report'
@ObjectModel.query.implementedBy: 'ABAP:ZCL_CE_SALES_REPORT'
@Metadata.allowExtensions: true
define custom entity zce_sales_report
{
  key Plant               : werks_d;
  key Product             : matnr;
  key CompanyCode         : abap.char(4);
  key FiscalYear          : gjahr;
      @ObjectModel.text.element: ['SalesDistrictName']
  key SalesDistrict       : abap.char(6);
      @EndUserText.label: 'Posting Date'
  key PostingDate         : abap.dats;
      @ObjectModel.text.element: ['ProductGroupName']
  key ProductGroup        : abap.char(9);
  key BaseUnit            : meins;

      ProfitCenter        : abap.char(10);
      ProductName         : abap.char(40);

      ProductGroupName    : abap.char(20);
      SalesDistrictName   : abap.char(20);

      @Semantics.quantity.unitOfMeasure: 'BaseUnit'
      @Aggregation.default: #SUM
      //      Quantity            : abap.dec(15,3);
      Quantity            : abap.quan(15,3);

      @Semantics.amount.currencyCode: 'CurrencyUSD'
      @Aggregation.default: #SUM
      RevenueUSD          : abap.curr(15,2);
      CurrencyUSD         : abap.cuky;

      @Semantics.amount.currencyCode: 'CompanyCodeCurrency'
      @Aggregation.default: #SUM
      RevenueVND          : abap.curr(15,2);

      @Semantics.amount.currencyCode: 'CompanyCodeCurrency'
      @Aggregation.default: #SUM
      COGSVND             : abap.curr(15,2);

      CompanyCodeCurrency : abap.cuky;
}
