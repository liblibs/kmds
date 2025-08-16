SELECT
  InvoLine.CashierID,
  Cashier.Name AS CashierName,
  Tender.ID AS TenderID,
  InvoLine.InputCode AS MOP,
  InvoLine.InvoHeadID AS InvoiceNumber,
  InvoLine.SetupDate,
  0 - [Price] AS Amount,
  Tender.GetCheckCashing,
  InvoLine.SalesDept AS CheckCashing,
  Tender.GetCheckNumber,
  InvoLine.SalesGroup AS CheckNumber,
  [FirstName] & " " & [LastName] AS CheckCashingName,
  Left([InputCode], 4) AS InputType
FROM
  (
    (
      InvoLine
      INNER JOIN Tender ON InvoLine.Code = Tender.ID
    )
    LEFT JOIN CheckCashing ON InvoLine.SalesDept = CheckCashing.ID
  )
  INNER JOIN Cashier ON InvoLine.CashierID = Cashier.ID
WHERE
  (
    (
      (Tender.ID)<> 1
    )
    AND (
      (0 - [Price])<> 0
    )
    AND (
      (
        Left([InputCode], 4)
      )<> "Drop"
      And (
        Left([InputCode], 4)
      )<> "Chan"
    )
    AND (
      (Tender.NonCashDetailReport)= Yes
    )
    AND (
      (InvoLine.TransType)= "T"
    )
  )
ORDER BY
  InvoLine.CashierID,
  Tender.ID,
  InvoLine.InputCode,
  0 - [Price],
  Tender.NonCashDetailReport;
