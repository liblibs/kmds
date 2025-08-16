SELECT
  Item.VendorID,
  Item.ID,
  Item.StockNumber,
  Item.Description,
  Size.Description AS SizeD,
  Item.Onhand,
  Item.Cost,
  IIf(
    IsNull([SalePrice]),
    [Price],
    [SalePrice]
  ) AS TodaysPrice,
  qryItemOnReserve.OnReserve,
  [onhand] + [TransferredToday] - [SoldTodayCr1] - [SoldTodayCr2] - [SoldTodayCr3] - [SoldTodayCr4] -(
    IIf(
      IsNull([OnReserve]),
      0,
      [OnReserve]
    )
  ) AS Avail,
  Item.OrderAt,
  Item.OrderTo
FROM
  (
    (
      Item
      INNER JOIN [Size] ON Item.SizeID = Size.ID
    )
    LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID
  )
  LEFT JOIN qryItemOnReserve ON Item.ID = qryItemOnReserve.ItemID
WHERE
  (
    (
      (Item.VendorID) Between [Forms]![PrintInventoryReport]![cboVendorIDBegin]
      And [Forms]![PrintInventoryReport]![cboVendorIDEnd]
    )
  );
