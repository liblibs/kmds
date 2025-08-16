SELECT
  InvoLine.Code,
  InvoLine.InputCode AS Transact,
  Sum(InvoLine.Quantity) AS SumOfQuantity,
  Sum(InvoLine.Price) AS SumOfPrice
FROM
  InvoLine
  INNER JOIN Tender ON InvoLine.Code = Tender.ID
GROUP BY
  InvoLine.Code,
  InvoLine.InputCode,
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
      (InvoLine.TransType)= "T"
    )
  )
ORDER BY
  InvoLine.Code;
