SELECT
  InvoLine.Code,
  InvoLine.InputCode AS Transact,
  Sum(InvoLine.Quantity) AS SumOfQuantity,
  Sum(InvoLine.Price) AS SumOfPrice
FROM
  InvoLine
  INNER JOIN Tender ON InvoLine.Code = Tender.ID
WHERE
  (
    (
      (InvoLine.TransType)= 'T'
    )
  )
GROUP BY
  InvoLine.Code,
  InvoLine.InputCode
HAVING
  (
    (
      (InvoLine.Code)= 1
    )
    AND (
      (InvoLine.InputCode)<> "Cash"
      And (InvoLine.InputCode)<> "Change"
    )
  )
ORDER BY
  InvoLine.Code;
