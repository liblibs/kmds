SELECT
  InvoLineAllRegisters.Code,
  InvoLineAllRegisters.InputCode AS Transact,
  Sum(InvoLineAllRegisters.Quantity) AS SumOfQuantity,
  Sum(InvoLineAllRegisters.Price) AS SumOfPrice,
  Cashier.ID,
  Cashier.Name
FROM
  Cashier
  INNER JOIN (
    InvoLineAllRegisters
    INNER JOIN Tender ON InvoLineAllRegisters.Code = Tender.ID
  ) ON Cashier.ID = InvoLineAllRegisters.CashierID
WHERE
  (
    (
      (InvoLineAllRegisters.TransType)= 'T'
    )
  )
GROUP BY
  InvoLineAllRegisters.Code,
  InvoLineAllRegisters.InputCode,
  Cashier.ID,
  Cashier.Name
HAVING
  (
    (
      (InvoLineAllRegisters.Code)= 1
    )
    AND (
      (InvoLineAllRegisters.InputCode)<> "Cash"
      And (InvoLineAllRegisters.InputCode)<> "Change"
    )
    AND (
      (Cashier.ID)= [parCashierNumber]
    )
  )
ORDER BY
  InvoLineAllRegisters.Code;
