SELECT
  InvoLine.Code,
  Min(InvoLine.SetupDate) AS MinOfSetupDate INTO ItemSetupDate
FROM
  InvoLine
GROUP BY
  InvoLine.Code;
