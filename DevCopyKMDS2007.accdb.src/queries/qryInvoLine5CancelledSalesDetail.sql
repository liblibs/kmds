SELECT
  qryInvoLine5CancelledSales.CashierID,
  Cashier.Name AS CashierName,
  qryInvoLine5CancelledSales.SetupDate,
  InvoLine5.CashSysID,
  InvoLine5.InvoHeadID,
  InvoLine5.Code,
  Item.Description,
  InvoLine5.InputCode,
  InvoLine5.Quantity,
  InvoLine5.Price,
  InvoLine5.ID
FROM
  (
    (
      qryInvoLine5CancelledSales
      INNER JOIN Cashier ON qryInvoLine5CancelledSales.CashierID = Cashier.ID
    )
    INNER JOIN InvoLine5 ON (
      qryInvoLine5CancelledSales.SetupDate = InvoLine5.SetupDate
    )
    AND (
      qryInvoLine5CancelledSales.CashSysID = InvoLine5.CashSysID
    )
    AND (
      qryInvoLine5CancelledSales.InvoHeadID = InvoLine5.InvoHeadID
    )
  )
  INNER JOIN Item ON InvoLine5.Code = Item.ID
WHERE
  (
    (
      (InvoLine5.ID)<= [maxid]
    )
  )
ORDER BY
  qryInvoLine5CancelledSales.CashierID,
  qryInvoLine5CancelledSales.SetupDate;
