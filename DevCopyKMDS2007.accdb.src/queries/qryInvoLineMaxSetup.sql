SELECT
  InvoLine.Code,
  Max(InvoLine.SetupDate) AS MaxOfSetupDate
FROM
  InvoLine
GROUP BY
  InvoLine.Code;
