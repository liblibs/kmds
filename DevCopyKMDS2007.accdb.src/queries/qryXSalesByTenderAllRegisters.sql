SELECT
  InvoLineAllRegisters.Code,
  Tender.Description,
  Sum(InvoLineAllRegisters.Quantity) AS SumOfQuantity,
  Sum(InvoLineAllRegisters.Price) AS SumOfPrice
FROM
  InvoLineAllRegisters
  INNER JOIN Tender ON InvoLineAllRegisters.Code = Tender.ID
WHERE
  (
    (
      (InvoLineAllRegisters.TransType)= 'T'
    )
  )
GROUP BY
  InvoLineAllRegisters.Code,
  Tender.Description
ORDER BY
  InvoLineAllRegisters.Code;
