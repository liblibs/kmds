INSERT INTO ItemsSoldToday (CashSysID, Code, SumOfQuantity)
SELECT
  InvoLine.CashSysID,
  InvoLine.Code,
  Sum(InvoLine.Quantity) AS SumOfQuantity
FROM
  InvoLine
GROUP BY
  InvoLine.CashSysID,
  InvoLine.Code;
