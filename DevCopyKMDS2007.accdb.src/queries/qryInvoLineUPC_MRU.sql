SELECT
  InvoLine.Code,
  InvoLine.InputCode,
  Max(InvoLine.SetupDate) AS MaxOfSetupDate
FROM
  InvoLine
WHERE
  (
    (
      (
        Len([inputcode])
      )> 7
    )
    AND (
      (
        IsNumeric([inputcode])
      )= True
    )
  )
GROUP BY
  InvoLine.Code,
  InvoLine.InputCode;
