SELECT
  involine.CashierID,
  Cashier.Name AS CashierName,
  involine.CashSysID AS CR,
  involine.InvoHeadID AS Invoice,
  Format(
    involine.SetupDate, "Short Time"
  ) AS [Time],
  ([Quantity]) AS TranCount,
  involine.Price AS Amount,
  involine.SalesDept AS CustomerID,
  [FirstName] & " " & [LastName] AS Customer,
  CheckCashing.LastName,
  involine.SalesGroup AS CheckNumber,
  involine.SetupDate
FROM
  Cashier
  INNER JOIN (
    (
      involine
      LEFT JOIN Tender ON involine.Code = Tender.ID
    )
    LEFT JOIN CheckCashing ON involine.SalesDept = CheckCashing.ID
  ) ON Cashier.ID = involine.CashierID
WHERE
  (
    (
      (involine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
    And (
      (involine.Code)= 5
    )
  )
ORDER BY
  involine.CashierID,
  involine.Price;
