SELECT
  InvoLine5.CashierID,
  Cashier.Name AS CashierName,
  InvoLine5.Code AS TenderID,
  Tender.Description,
  Left([InputCode], 4) AS CashORLoan,
  InvoLine5.InputCode AS Transact,
  InvoLine5.CashSysID AS CR,
  InvoLine5.InvoHeadID AS Invoice,
  Format(
    InvoLine5.SetupDate, "Short Time"
  ) AS [Time],
  Sum(
    (1)
  ) AS TranCount,
  Sum(InvoLine5.Price) AS Amount,
  Sum(
    IIf(
      Left([InputCode], 4)= "Cash",
      0 - [Price],
      0
    )
  ) AS Loan,
  Sum(
    IIf(
      Left([InputCode], 4)= "Drop",
      [Price],
      0
    )
  ) AS [Drop]
FROM
  Cashier
  INNER JOIN (
    InvoLine5
    INNER JOIN Tender ON InvoLine5.Code = Tender.ID
  ) ON Cashier.ID = InvoLine5.CashierID
GROUP BY
  InvoLine5.CashierID,
  Cashier.Name,
  InvoLine5.Code,
  Tender.Description,
  Left([InputCode], 4),
  InvoLine5.InputCode,
  InvoLine5.CashSysID,
  InvoLine5.InvoHeadID,
  Format(
    InvoLine5.SetupDate, "Short Time"
  ),
  InvoLine5.TransType
HAVING
  (
    (
      (
        Left([InputCode], 4)
      )= "Cash"
      Or (
        Left([InputCode], 4)
      )= "Drop"
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
  InvoLine5.CashierID,
  InvoLine5.Code,
  Left([InputCode], 4),
  InvoLine5.InvoHeadID,
  Format(
    InvoLine5.SetupDate, "Short Time"
  ),
  Sum(InvoLine5.Price);
