SELECT
  InvoLine.Code,
  Sum(
    IIf(
      [Transtype] = "S"
      Or [Transtype] = "C",
      [Quantity], 0
    )
  ) AS Sq
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.SetupDate)> Now()-365
    )
  )
GROUP BY
  InvoLine.Code;
