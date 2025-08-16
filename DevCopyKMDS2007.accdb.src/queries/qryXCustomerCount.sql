SELECT
  DISTINCT InvoLine.CashSysID,
  InvoLine.CashierID,
  InvoLine.InvoHeadID,
  Sum(1) AS [count],
  Sum(InvoLine.Price) AS Sales
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.TransType)= "S"
    )
  )
GROUP BY
  InvoLine.CashSysID,
  InvoLine.CashierID,
  InvoLine.InvoHeadID;
