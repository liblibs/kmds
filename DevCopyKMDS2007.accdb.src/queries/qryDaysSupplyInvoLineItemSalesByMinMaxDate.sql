SELECT
  InvoLine.Code,
  Sum(InvoLine.Quantity) AS SumOfQuantity,
  Sum(InvoLine.Price) AS SumOfPrice,
  Min(InvoLine.SetupDate) AS MinOfSetupDate,
  Max(InvoLine.SetupDate) AS MaxOfSetupDate
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.TransType)= "S"
      Or (InvoLine.TransType)= "C"
    )
  )
GROUP BY
  InvoLine.Code
ORDER BY
  Min(InvoLine.SetupDate);
