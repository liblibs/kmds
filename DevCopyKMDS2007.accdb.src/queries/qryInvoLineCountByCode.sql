PARAMETERS parCode Decimal;
SELECT
  InvoLine.Code,
  Sum(1) AS [Count]
FROM
  InvoLine
GROUP BY
  InvoLine.Code
HAVING
  (
    (
      (InvoLine.Code)= [parCode]
    )
  );
