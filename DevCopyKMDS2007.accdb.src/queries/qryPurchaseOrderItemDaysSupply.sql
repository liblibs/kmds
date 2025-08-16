PARAMETERS parSupplyDays Short;
SELECT
  qryPurchaseOrderItemDaysSupplyADS.VendorID,
  Vendor.Name AS VendorName,
  Item.SalesDepartmentID,
  qryPurchaseOrderItemDaysSupplyADS.ItemId,
  Item.Description,
  Size.Description AS SizeDesc,
  qryPurchaseOrderItemDaysSupplyADS.Sold,
  qryPurchaseOrderItemDaysSupplyADS.SoldDays,
  ([Sold] / [SoldDays]) AS AverageDailySales,
  [parSupplyDays] AS SupplyDays,
  CInt(
    ([Sold] / [SoldDays])* [parSupplyDays]
  ) AS SDOrderTo,
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
  (
    [Onhand] + [Ordered] - [Reserved] - [SoldTodayCr1] - [SoldTodayCr2] - [SoldTodayCr3] - [SoldTodayCr4] + [TransferredToday]
  ) AS Avail,
  (
    (
      CInt(
        ([Sold] / [SoldDays])* [parSupplyDays]
      )
    )- [Avail]
  ) AS SOQRaw,
  Item.UnitsPerOrder AS OrderMultiple,
  CInt(
    (
      ([SOQRaw])/ [UnitsPerOrder]
    )
  )* [UnitsPerOrder] AS SOQ,
  qryPurchaseOrderItemDaysSupplyADS.DateStart,
  qryPurchaseOrderItemDaysSupplyADS.DateEnd,
  Item.UnitsPerCase,
  IIf(
    [soq] Mod [UnitsPerCase] <> 0, True,
    False
  ) AS Repack,
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
  Item.OrderAt AS ItemOrderAt,
  Item.OrderTo
FROM
  qryPurchaseOrderItemDaysSupplyADS
  INNER JOIN (
    (
      (
        (
          (
            Item
            INNER JOIN Vendor ON Item.VendorID = Vendor.ID
          )
          LEFT JOIN qryItemOnReserve ON Item.ID = qryItemOnReserve.ItemID
        )
        LEFT JOIN qryItemOnOrder ON Item.ID = qryItemOnOrder.ItemID
      )
      LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID
    )
    INNER JOIN [Size] ON Item.SizeID = Size.ID
  ) ON qryPurchaseOrderItemDaysSupplyADS.ItemId = Item.ID
WHERE
  (
    (
      (
        qryPurchaseOrderItemDaysSupplyADS.Sold
      )<> 0
    )
    AND (
      (Item.Onhand) Is Not Null
    )
  )
ORDER BY
  Item.SalesDepartmentID,
  Item.Description;
