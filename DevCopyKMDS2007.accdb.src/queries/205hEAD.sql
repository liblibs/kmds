SELECT
  InvoLine.InvoHeadID,
  InvoLine.ShiftID,
  InvoLine.CashSysID,
  InvoLine.SetupDate,
  Sum(1) AS [cOUNT]
FROM
  InvoLine
GROUP BY
  InvoLine.InvoHeadID,
  InvoLine.ShiftID,
  InvoLine.CashSysID,
  InvoLine.SetupDate
HAVING
  (
    (
      (
        Sum(1)
      )= 1
    )
  );
