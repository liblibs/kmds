SELECT
  qryEOSalesReturnsByCashierByItem.CashierID,
  Sum(
    qryEOSalesReturnsByCashierByItem.SumOfQuantity
  ) AS ReturnQty,
  Sum(
    qryEOSalesReturnsByCashierByItem.SumOfPrice
  ) AS ReturnDollars
FROM
  qryEOSalesReturnsByCashierByItem
GROUP BY
  qryEOSalesReturnsByCashierByItem.CashierID;
