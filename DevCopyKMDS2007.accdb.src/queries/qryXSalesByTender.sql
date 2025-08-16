SELECT
  InvoLine.Code,
  Tender.Description,
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
  Tender.Description
ORDER BY
  InvoLine.Code;
