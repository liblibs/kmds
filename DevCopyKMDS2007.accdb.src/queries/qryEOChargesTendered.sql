SELECT
  InvoLine5.CashSysID AS CR,
  InvoLine5.CashierID,
  Cashier.Name AS CashierName,
  InvoLine5.InvoHeadID AS Invoice,
  InvoLine5.SalesDept,
  Format(
    InvoLine5.SetupDate, "Short Time"
  ) AS [Time],
  (0 - [Quantity]) AS TranCount,
  0 - [Price] AS Amount,
  InvoLine5.SalesGroup
FROM
  Cashier
  INNER JOIN (
    InvoLine5
    LEFT JOIN Tender ON InvoLine5.Code = Tender.ID
  ) ON Cashier.ID = InvoLine5.CashierID
WHERE
  (
    (
      (InvoLine5.Code)= 3
    )
  )
ORDER BY
  InvoLine5.CashSysID,
  InvoLine5.CashierID,
  InvoLine5.InvoHeadID;
