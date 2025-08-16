SELECT
  InvoLine.InvoHeadID,
  InvoLine.CashSysID,
  InvoLine.SetupDate,
  Sum(InvoLine.Price) AS SumOfPrice
FROM
  Sysfile
  INNER JOIN InvoLine ON Sysfile.CouponItemID = InvoLine.Code
GROUP BY
  InvoLine.InvoHeadID,
  InvoLine.CashSysID,
  InvoLine.SetupDate
ORDER BY
  InvoLine.SetupDate;
