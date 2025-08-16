SELECT
  InvoHead5.CashierID,
  Cashier.Name AS CashierName,
  InvoLine5.Code AS TenderID,
  InvoLine5.InputCode AS Transact,
  InvoLine5.CashSysID AS CR,
  InvoLine5.InvoHeadID AS Invoice,
  Format(
    Invohead5.SetupDate, "Short Time"
  ) AS [Time],
  Sum(
    (1)
  ) AS TranCount,
  Sum(0 - [Price]) AS Amount
FROM
  (
    InvoHead5
    INNER JOIN Cashier ON InvoHead5.CashierID = Cashier.ID
  )
  INNER JOIN (
    InvoLine5
    INNER JOIN Tender ON InvoLine5.Code = Tender.ID
  ) ON (
    InvoHead5.CashSysID = InvoLine5.CashSysID
  )
  AND (
    InvoHead5.ID = InvoLine5.InvoHeadID
  )
GROUP BY
  InvoHead5.CashierID,
  Cashier.Name,
  InvoLine5.Code,
  InvoLine5.InputCode,
  InvoLine5.CashSysID,
  InvoLine5.InvoHeadID,
  Format(
    Invohead5.SetupDate, "Short Time"
  ),
  InvoLine5.TransType
HAVING
  (
    (
      (InvoLine5.Code)= 1
    )
    AND (
      (InvoLine5.InputCode)<> "Cash"
      And (InvoLine5.InputCode)<> "Change"
    )
    AND (
      (InvoLine5.TransType)= 'T'
    )
  )
ORDER BY
  InvoHead5.CashierID,
  InvoLine5.InputCode,
  Sum(0 - [Price]);
