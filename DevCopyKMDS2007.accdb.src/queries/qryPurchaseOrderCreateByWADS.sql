SELECT
  Item.ID,
  Item.Description,
  Item.AverageDailySales,
  DaysSupplyGroupLevels.DaysSupply,
  CLng(
    IIf(
      IsNull([AverageDailySales]),
      0,
      [AverageDailySales]
    )* [DaysSupply]
  ) AS OrderMax,
  Item.Onhand,
  [OrderMax] - [Onhand] AS RSOQ,
  Item.VendorID,
  (
    CLng([RSOQ] / [UnitsPerOrder])* [UnitsPerOrder]
  ) AS SOQ,
  Item.UnitsPerOrder
FROM
  Item
  INNER JOIN DaysSupplyGroupLevels ON Item.DaysSupplyGroupID = DaysSupplyGroupLevels.GroupLevel
WHERE
  (
    (
      (Item.VendorID)= 1
    )
  )
ORDER BY
  Item.ID;
