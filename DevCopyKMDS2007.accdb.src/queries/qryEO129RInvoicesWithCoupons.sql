SELECT
  InvoLine5.InvoHeadID,
  InvoLine5.CashSysID,
  InvoLine5.SetupDate,
  Sysfile.CouponItemID,
  Item.Description AS CouponDescription,
  Sum(InvoLine5.Price) AS SumOfPrice
FROM
  (
    Sysfile
    INNER JOIN InvoLine5 ON Sysfile.CouponItemID = InvoLine5.Code
  )
  INNER JOIN Item ON InvoLine5.Code = Item.ID
GROUP BY
  InvoLine5.InvoHeadID,
  InvoLine5.CashSysID,
  InvoLine5.SetupDate,
  Sysfile.CouponItemID,
  Item.Description;
