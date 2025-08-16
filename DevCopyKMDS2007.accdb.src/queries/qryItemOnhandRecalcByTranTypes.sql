SELECT
  InvoLine.Code,
  Sum(
    IIf(
      [TransType] = "S"
      Or [TransType] = "C",
      0 - [Quantity], [Quantity]
    )
  ) AS Movement
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.TransType)= "S"
      Or (InvoLine.TransType)= "C"
      Or (InvoLine.TransType)= "P"
      Or (InvoLine.TransType)= "X"
      Or (InvoLine.TransType)= "A"
    )
    And (
      (InvoLine.SetupDate)> CDate(
        Forms!ReportsExplorer!txtDateRangeFrom
      )
    )
  )
GROUP BY
  InvoLine.Code;
