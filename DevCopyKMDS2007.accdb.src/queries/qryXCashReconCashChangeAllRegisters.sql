SELECT
  InvoLineAllRegisters.Code,
  InvoLineAllRegisters.InputCode AS Transact,
  Sum(InvoLineAllRegisters.Quantity) AS SumOfQuantity,
  Sum(InvoLineAllRegisters.Price) AS SumOfPrice
FROM
  InvoLineAllRegisters
  INNER JOIN Tender ON InvoLineAllRegisters.Code = Tender.ID
WHERE
  (
    (
      (InvoLineAllRegisters.TransType)= 'T'
      Or (InvoLineAllRegisters.TransType)= 'T'
    )
  )
GROUP BY
  InvoLineAllRegisters.Code,
  InvoLineAllRegisters.InputCode
HAVING
  (
    (
      (InvoLineAllRegisters.Code)= 1
    )
    AND (
      (InvoLineAllRegisters.InputCode)= "Cash"
    )
  )
  OR (
    (
      (InvoLineAllRegisters.Code)= 1
    )
    AND (
      (InvoLineAllRegisters.InputCode)= "Change"
    )
  )
ORDER BY
  InvoLineAllRegisters.Code;
