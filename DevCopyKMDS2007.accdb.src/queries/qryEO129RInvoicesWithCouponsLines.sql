SELECT
  qryEO129RInvoicesWithCoupons.CouponItemID,
  qryEO129RInvoicesWithCoupons.CouponDescription,
  InvoLine5.InvoHeadID,
  InvoLine5.CashSysID,
  InvoLine5.ID,
  InvoLine5.SetupDate,
  InvoLine5.CashierID,
  Cashier.Name,
  InvoLine5.Code,
  Item.Description,
  InvoLine5.Quantity,
  InvoLine5.Price,
  InvoLine5.InputCode,
  InvoLine5.TransType,
  IIf(
    [code] = [CouponItemID], [InvoLine5].[Price],
    0
  ) AS CouponTotal,
  IIf([code] = [CouponItemID], "*", "") AS IsCoupon
FROM
  (
    (
      qryEO129RInvoicesWithCoupons
      INNER JOIN InvoLine5 ON (
        qryEO129RInvoicesWithCoupons.SetupDate = InvoLine5.SetupDate
      )
      AND (
        qryEO129RInvoicesWithCoupons.CashSysID = InvoLine5.CashSysID
      )
      AND (
        qryEO129RInvoicesWithCoupons.InvoHeadID = InvoLine5.InvoHeadID
      )
    )
    LEFT JOIN Item ON InvoLine5.Code = Item.ID
  )
  INNER JOIN Cashier ON InvoLine5.CashierID = Cashier.ID;
