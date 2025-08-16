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
      (
        Format(
          [InvoLine].[SetupDate], "yyyymm"
        )
      )>= Format(
        (
          Now()-365
        ),
        "yyyymm"
      )
      And (
        Format(
          [InvoLine].[SetupDate], "yyyymm"
        )
      )< Format(
        (
          Now()
        ),
        "yyyymm"
      )
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
