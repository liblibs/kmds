SELECT
  InvoLine.Code,
  InvoLine.SetupDate,
  InvoLine.InvoHeadID,
  InvoLine.ID,
  InvoLine.CashSysID,
  InvoLine.CashierID,
  InvoLine.ShiftID,
  InvoLine.SalesDept,
  InvoLine.InputCode
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.Code)= 92
    )
  )
ORDER BY
  InvoLine.SetupDate;
