SELECT
  Format([SetupDate], "yyyymmdd") AS BusinessDATE,
  InvoLine.Code,
  Sum(1) AS [COUNT],
  Sum(InvoLine.Price) AS SumOfPrice
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.TransType)= "T"
    )
    AND (
      (InvoLine.SetupDate) Between #1/1/2013#
      And #12/31/2013#
    )
  )
GROUP BY
  Format([SetupDate], "yyyymmdd"),
  InvoLine.Code
ORDER BY
  Format([SetupDate], "yyyymmdd"),
  InvoLine.Code;
