SELECT
  InvoLine.Code,
  Sum(InvoLine.Quantity) AS Sold INTO DaysSupplyItemSales
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.TransType)= "S"
      Or (InvoLine.TransType)= "C"
    )
    AND (
      (InvoLine.SetupDate)> Now()-365
    )
  )
GROUP BY
  InvoLine.Code
ORDER BY
  InvoLine.Code;
