SELECT
  InvoLine.Code,
  InvoLine.InputCode AS Transact,
  Sum(InvoLine.Quantity) AS SumOfQuantity,
  Sum(InvoLine.Price) AS SumOfPrice,
  Cashier.ID,
  Cashier.Name
FROM
  (
    (
      InvoLine
      INNER JOIN InvoHead ON (
        InvoLine.CashSysID = InvoHead.CashSysID
      )
      AND (
        InvoLine.InvoHeadID = InvoHead.ID
      )
    )
    INNER JOIN Cashier ON InvoHead.CashierID = Cashier.ID
  )
  INNER JOIN Tender ON InvoLine.Code = Tender.ID
WHERE
  (
    (
      (InvoLine.TransType)= 'T'
    )
  )
GROUP BY
  InvoLine.Code,
  InvoLine.InputCode,
  Cashier.ID,
  Cashier.Name
HAVING
  (
    (
      (InvoLine.Code)= 1
    )
    AND (
      (InvoLine.InputCode)<> "Cash"
      And (InvoLine.InputCode)<> "Change"
    )
    AND (
      (Cashier.ID)= [parCashierNumber]
    )
  )
ORDER BY
  InvoLine.Code;
