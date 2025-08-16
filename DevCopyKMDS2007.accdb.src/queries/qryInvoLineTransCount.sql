SELECT
  InvoLine.Code,
  Sum(1) AS [count]
FROM
  InvoLine
GROUP BY
  InvoLine.Code;
