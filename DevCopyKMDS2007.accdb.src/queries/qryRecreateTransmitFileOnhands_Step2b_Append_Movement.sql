INSERT INTO RecreateInventoryTransmit (Code, Movement)
SELECT
  InvoLine.Code,
  Sum(
    IIf(
      [TransType] = "S"
      Or [TransType] = "C",
      [Quantity], 0 - [Quantity]
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
      (InvoLine.SetupDate)> Forms!ReportsExplorer!txtDateRangeTo
    )
  )
GROUP BY
  InvoLine.Code;
