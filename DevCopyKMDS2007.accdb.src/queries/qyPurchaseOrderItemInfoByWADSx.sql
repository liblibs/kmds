SELECT
  Vendor.ID AS VendorID,
  Vendor.Name,
  Item.ID AS ItemID,
  Item.Description,
  Item.SizeID,
  Item.Onhand,
  IIf(
    IsNull([OnOrder]),
    0,
    [OnOrder]
  ) AS Ordered,
  IIf(
    IsNull([OnReserve]),
    0,
    [OnReserve]
  ) AS Reserved,
  [Onhand] + [Ordered] - [Reserved] - [SoldTodayCr1] - [SoldTodayCr2] - [SoldTodayCr3] - [SoldTodayCr4] + [TransferredToday] AS Avail,
  Item.AverageDailySales,
  AverageDaySupplyLevels.DaysSupply,
  CLng(
    IIf(
      IsNull([AverageDailySales]),
      0,
      [AverageDailySales]
    )* [DaysSupply]
  ) AS OrderMax,
  ([OrderMax] - [Avail] + 0.01) AS SOQRaw,
  IIf(
    [UnitsPerOrder] < 1, 1, [UnitsPerOrder]
  ) AS OrderMultiple,
  (
    CInt([SOQRaw] / [OrderMultiple])* [OrderMultiple]
  ) AS SOQ,
  Item.UnitsPerCase,
  Item.Price,
  IIf(
    IsNull([SalePrice]),
    Item.Price,
    [SalePrice]
  ) AS SPrice,
  IIf(
    IsNull([SalePrice]),
    " ",
    "*"
  ) AS SPriceStar,
  [SOQ] * [SPrice] AS ExtPrice,
  Item.Cost,
  [SOQ] * [Cost] AS ExtCost,
  Item.SalesDepartmentID,
  IIf(
    [soq] Mod [UnitsPerCase] <> 0, True,
    False
  ) AS Repack INTO PurchaseOrderItemWADS
FROM
  (
    (
      (
        (
          Vendor
          INNER JOIN Item ON Vendor.ID = Item.VendorID
        )
        LEFT JOIN qryItemOnOrder ON Item.ID = qryItemOnOrder.ItemID
      )
      LEFT JOIN qryItemOnReserve ON Item.ID = qryItemOnReserve.ItemID
    )
    LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID
  )
  INNER JOIN AverageDaySupplyLevels ON Item.DaysSupplyGroupID = AverageDaySupplyLevels.ID
WHERE
  (
    (
      (Vendor.ID)= [parVendorID]
    )
  )
ORDER BY
  Item.ID;
