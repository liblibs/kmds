SELECT
  Item.SalesDepartmentID,
  SalesDepartment.Description AS DepartmentDesc,
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
  [onhand] + [TransferredToday] - [SoldTodayCr1] - [SoldTodayCr2] - [SoldTodayCr3] - [SoldTodayCr4] AS OnTheShelf
FROM
  SalesDepartment
  INNER JOIN (
    (
      (
        Item
        INNER JOIN [Size] ON Item.SizeID = Size.ID
      )
      LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID
    )
    LEFT JOIN qryItemOnReserve ON Item.ID = qryItemOnReserve.ItemID
  ) ON SalesDepartment.ID = Item.SalesDepartmentID
WHERE
  (
    (
      (Item.SalesDepartmentID) Between [Forms]![PrintInventoryReport]![cboSalesDepartmentIDBegin]
      And [Forms]![PrintInventoryReport]![cboSalesDepartmentIDEnd]
    )
  )
ORDER BY
  Item.SalesDepartmentID;
