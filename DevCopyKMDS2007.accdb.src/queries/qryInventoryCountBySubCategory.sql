SELECT
  Item.SalesDepartmentID,
  SubCategory.ID AS SubCategoryID,
  SubCategory.Description AS SubCatDesc,
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
  SubCategory
  INNER JOIN (
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
  ) ON SubCategory.ID = Item.SubCategoryID
WHERE
  (
    (
      (Item.SalesDepartmentID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (SubCategory.Description) Between Forms!PrintInventoryReport!cboCategoryIDBegin
      And Forms!PrintInventoryReport!cboCategoryIDEnd
    )
  )
ORDER BY
  Item.SalesDepartmentID,
  SubCategory.Description,
  Item.ID;
