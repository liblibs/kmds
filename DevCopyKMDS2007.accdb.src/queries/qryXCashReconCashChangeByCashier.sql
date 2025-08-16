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
      InvoHead
      INNER JOIN InvoLine ON (
        InvoHead.ID = InvoLine.InvoHeadID
      )
      AND (
        InvoHead.CashSysID = InvoLine.CashSysID
      )
    )
    INNER JOIN Cashier ON InvoHead.CashierID = Cashier.ID
  )
  INNER JOIN Tender ON InvoLine.Code = Tender.ID
GROUP BY
  InvoLine.Code,
  InvoLine.InputCode,
  Cashier.ID,
  Cashier.Name,
  InvoLine.TransType
HAVING
  (
    (
      (InvoLine.Code)= 1
    )
    AND (
      (InvoLine.InputCode)= "Cash"
    )
    AND (
      (Cashier.ID)= [parCashierNumber]
    )
    AND (
      (InvoLine.TransType)= "T"
    )
  )
  OR (
    (
      (InvoLine.Code)= 1
    )
    AND (
      (InvoLine.InputCode)= "Change"
    )
    AND (
      (Cashier.ID)= [parCashierNumber]
    )
    AND (
      (InvoLine.TransType)= "T"
    )
  )
ORDER BY
  InvoLine.Code;
