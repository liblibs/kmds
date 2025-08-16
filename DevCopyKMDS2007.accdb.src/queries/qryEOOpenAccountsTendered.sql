SELECT
  InvoLine5.CashierID,
  Cashier.Name AS CashierName,
  InvoLine5.CashSysID AS CR,
  InvoLine5.InvoHeadID AS Invoice,
  Format(
    InvoLine5.SetupDate, "Short Time"
  ) AS [Time],
  ([Quantity]) AS TranCount,
  InvoLine5.Price AS Amount,
  InvoLine5.SalesDept AS CustomerID,
  [FirstName] & " " & [LastName] AS Customer,
  CheckCashing.LastName,
  InvoLine5.SalesGroup AS CheckNumber
FROM
  Cashier
  INNER JOIN (
    (
      InvoLine5
      LEFT JOIN Tender ON InvoLine5.Code = Tender.ID
    )
    LEFT JOIN CheckCashing ON InvoLine5.SalesDept = CheckCashing.ID
  ) ON Cashier.ID = InvoLine5.CashierID
WHERE
  (
    (
      (InvoLine5.Code)= 5
    )
  )
ORDER BY
  InvoLine5.CashierID,
  InvoLine5.Price;
