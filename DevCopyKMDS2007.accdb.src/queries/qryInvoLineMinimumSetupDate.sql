SELECT
  InvoLine.Code,
  Min(InvoLine.SetupDate) AS MinOfSetupDate
FROM
  InvoLine
  INNER JOIN Item ON InvoLine.Code = Item.ID
GROUP BY
  InvoLine.Code;
