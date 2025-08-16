SELECT
  qryInvoLine5SalesByCashierSummary.ID,
  qryInvoLine5SalesByCashierSummary.Name,
  qryInvoLine5SalesByCashierSummary.TotalTransactions,
  qryInvoLine5SalesByCashierSummary.TotalSales,
  qryInvoLine5SalesByCashierSummary.AvgSale,
  qryInvoLine5SalesByCashierSummary.TotalNoSales,
  qryInvoLine5SalesByCashierSummary.TotalCancelledSales,
  qryInvoLine5SalesByCashierSummary.TotalIDScanCount,
  IIf(
    IsNull([ReturnQty]),
    0,
    [ReturnQty]
  ) AS ReturnItems,
  IIf(
    IsNull([ReturnDollars]),
    0,
    [ReturnDollars]
  ) AS ReturnAmount
FROM
  qryInvoLine5SalesByCashierSummary
  LEFT JOIN qrySalesReturnsByCashier ON qryInvoLine5SalesByCashierSummary.ID = qrySalesReturnsByCashier.CashierID;
