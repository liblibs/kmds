SELECT
  InvoLine.Code,
  Max(InvoLine.SetupDate) AS MaxOfSetupDate
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.TransType)= "Y"
    )
  )
GROUP BY
  InvoLine.Code;
