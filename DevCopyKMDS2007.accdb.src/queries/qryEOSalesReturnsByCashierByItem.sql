SELECT
  InvoLine5.CashierID,
  Cashier.Name AS CashierName,
  InvoLine5.TransType,
  InvoLine5.SetupDate,
  InvoLine5.CashSysID,
  InvoLine5.InvoHeadID,
  InvoLine5.Code,
  Item.Description,
  Sum(InvoLine5.Quantity) AS SumOfQuantity,
  Sum(InvoLine5.Price) AS SumOfPrice
FROM
  Sysfile,
  Cashier
  INNER JOIN (
    InvoLine5
    INNER JOIN Item ON InvoLine5.Code = Item.ID
  ) ON Cashier.ID = InvoLine5.CashierID
WHERE
  (
    (
      (InvoLine5.TransType)= "S"
    )
    And (
      (InvoLine5.Code)<> Sysfile.CouponItemID
      And (InvoLine5.Code)<> Sysfile.DiscountItemID
      And (InvoLine5.Code)<> Sysfile.LotteryPaidOutItemID
    )
  )
GROUP BY
  InvoLine5.CashierID,
  Cashier.Name,
  InvoLine5.TransType,
  InvoLine5.SetupDate,
  InvoLine5.CashSysID,
  InvoLine5.InvoHeadID,
  InvoLine5.Code,
  Item.Description
HAVING
  (
    (
      (
        Sum(InvoLine5.Quantity)
      )< 0
    )
    AND (
      (InvoLine5.TransType)= "S"
    )
  )
ORDER BY
  InvoLine5.CashierID,
  InvoLine5.SetupDate;
