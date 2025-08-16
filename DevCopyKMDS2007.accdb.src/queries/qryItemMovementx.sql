SELECT
  InvoLine.Code,
  InvoLine.TransType,
  Sum(InvoLine.Quantity) AS SumOfQuantity,
  Sum(InvoLine.Price) AS SumOfPrice,
  Format(InvoLine.SetupDate, "yyyy") AS YYYY,
  Format(InvoLine.SetupDate, "mm") AS MM
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.SetupDate) Between #3/1/2004#
      And #3/1/2005#
    )
  )
GROUP BY
  InvoLine.Code,
  InvoLine.TransType,
  Format(InvoLine.SetupDate, "yyyy"),
  Format(InvoLine.SetupDate, "mm")
ORDER BY
  InvoLine.Code,
  Format(InvoLine.SetupDate, "yyyy"),
  Format(InvoLine.SetupDate, "mm");
