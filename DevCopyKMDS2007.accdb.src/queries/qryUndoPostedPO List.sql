SELECT
  InvoLine.InvoHeadID AS InvoiceHeadID,
  InvoLine.TransType,
  InvoLine.InputCode,
  Sum(InvoLine.Quantity) AS SumOfQuantity,
  Sum(InvoLine.Price) AS SumOfPrice,
  Format([SetupDate], "yyyy-mm-dd") AS [Date]
FROM
  InvoLine
GROUP BY
  InvoLine.InvoHeadID,
  InvoLine.TransType,
  InvoLine.InputCode,
  Format([SetupDate], "yyyy-mm-dd")
HAVING
  (
    (
      (InvoLine.TransType)= "P"
    )
    AND (
      (
        Format([SetupDate], "yyyy-mm-dd")
      )> Format(
        Now()-30,
        "yyyy-mm-dd"
      )
    )
  )
ORDER BY
  Format([SetupDate], "yyyy-mm-dd") DESC;
