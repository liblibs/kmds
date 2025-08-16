SELECT
  Involine.CashierID,
  Cashier.Name AS CashierName,
  Involine.TransType,
  Involine.SetupDate,
  Involine.CashSysID,
  Involine.InvoHeadID,
  Involine.Code,
  Item.Description,
  Sum(Involine.Quantity) AS SumOfQuantity,
  Sum(Involine.Price) AS SumOfPrice
FROM
  Sysfile,
  Cashier
  INNER JOIN (
    Involine
    INNER JOIN Item ON Involine.Code = Item.ID
  ) ON Cashier.ID = Involine.CashierID
WHERE
  (
    (
      (Involine.TransType)= "S"
    )
    And (
      (Involine.Code)<> Sysfile.CouponItemID
      And (Involine.Code)<> Sysfile.DiscountItemID
      And (Involine.Code)<> Sysfile.LotteryPaidOutItemID
    )
  )
GROUP BY
  Involine.CashierID,
  Cashier.Name,
  Involine.TransType,
  Involine.SetupDate,
  Involine.CashSysID,
  Involine.InvoHeadID,
  Involine.Code,
  Item.Description
HAVING
  (
    (
      (Involine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
    And (
      (
        Sum(Involine.Quantity)
      )< 0
    )
    And (
      (Involine.TransType)= "S"
    )
  )
ORDER BY
  Involine.CashierID,
  Involine.SetupDate;
