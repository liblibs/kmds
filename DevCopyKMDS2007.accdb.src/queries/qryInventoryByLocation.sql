SELECT
  Item.Location,
  Item.ID,
  Item.Description,
  Size.Description AS SizeD,
  Item.Onhand,
  Item.Cost,
  Item.Price,
  qryItemsSalePriceNow.SalePrice,
  IIf(
    IsNull([SalePrice]),
    [Price],
    [SalePrice]
  ) AS TodaysPrice,
  qryItemOnReserve.OnReserve,
  [onhand] + [TransferredToday] - [SoldTodayCr1] - [SoldTodayCr2] - [SoldTodayCr3] - [SoldTodayCr4] AS OnTheShelf,
  Item.VendorID,
  Vendor.Name AS VendorName,
  Vendor.POMethodDefault
FROM
  (
    (
      (
        Item
        INNER JOIN [Size] ON Item.SizeID = Size.ID
      )
      LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID
    )
    LEFT JOIN qryItemOnReserve ON Item.ID = qryItemOnReserve.ItemID
  )
  INNER JOIN Vendor ON Item.VendorID = Vendor.ID
WHERE
  (
    (
      (Item.Location) Between Forms!PrintInventoryReport!cboLocationBegin
      And Forms!PrintInventoryReport!cboLocationEnd
    )
    And (
      (Item.VendorID) Between Forms!PrintInventoryReport!cboVendorIDBegin
      And Forms!PrintInventoryReport!cboVendorIDEnd
    )
    And (
      (Vendor.POMethodDefault)<>-1
    )
  )
ORDER BY
  Item.Location,
  Item.Description;
