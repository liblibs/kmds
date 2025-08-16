SELECT
  Cashier.ID,
  Cashier.Name,
  Sum(
    qryInvoLine5SalesByCashierByInvoice.Transactions
  ) AS TotalTransactions,
  Sum(
    qryInvoLine5SalesByCashierByInvoice.Sales
  ) AS TotalSales,
  [TotalSales] / [TotalTransactions] AS AvgSale,
  Sum(
    qryInvoLine5SalesByCashierByInvoice.NoSales
  ) AS TotalNoSales,
  Sum(
    qryInvoLine5SalesByCashierByInvoice.CancelledSales
  ) AS TotalCancelledSales,
  Sum(
    qryInvoLine5SalesByCashierByInvoice.IDScanCount
  ) AS TotalIDScanCount
FROM
  Cashier
  INNER JOIN qryInvoLine5SalesByCashierByInvoice ON Cashier.ID = qryInvoLine5SalesByCashierByInvoice.CashierID
GROUP BY
  Cashier.ID,
  Cashier.Name;
