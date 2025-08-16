SELECT
  InvoLine.CashSysID,
  InvoLine.CashierID,
  Sum(InvoLine.Code) AS NoSales
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.Code)= 1
    )
  )
GROUP BY
  InvoLine.CashSysID,
  InvoLine.CashierID;
