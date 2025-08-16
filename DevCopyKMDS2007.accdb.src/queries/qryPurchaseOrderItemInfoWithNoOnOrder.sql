SELECT
  Vendor.ID AS VendorID,
  Vendor.Name,
  Item.ID AS ItemID,
  Item.Description,
  Item.SizeID,
  Item.Onhand,
  0 AS Ordered,
  IIf(
    IsNull([OnReserve]),
    0,
    [OnReserve]
  ) AS Reserved,
  [Onhand] + [Ordered] - [Reserved] - [SoldTodayCr1] - [SoldTodayCr2] - [SoldTodayCr3] - [SoldTodayCr4] + [TransferredToday] AS Avail,
  Item.OrderAt,
  Item.OrderTo,
  [OrderTo] - [Avail] AS SOQRaw,
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
  Item.SalesDepartmentID
FROM
  (
    (
      (
        Vendor
        INNER JOIN Item ON Vendor.ID = Item.VendorID
      )
      LEFT JOIN qryItemOnReserve ON Item.ID = qryItemOnReserve.ItemID
    )
    LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID
  )
  INNER JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
WHERE
  (
    (
      (SubCategory.Description)<> "OBSOLETE"
    )
  )
ORDER BY
  Item.ID;
