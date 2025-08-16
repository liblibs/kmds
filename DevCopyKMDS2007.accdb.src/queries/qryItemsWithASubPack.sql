SELECT
  Item.SalesDepartmentID AS CartonSD,
  SalesDepartment.Description AS CartonSDDesc,
  Item.ID AS CartonID,
  Item.Description AS CartonDesc,
  Item.SubPack AS CartonSUBUpc,
  Item.UnitsPerPack AS CartonPack,
  Item.Price AS CartonPrice,
  Item_1.SalesDepartmentID AS PackSD,
  SalesDepartment_1.Description AS PackSDDesc,
  Item_1.ID AS PackID,
  Item_1.Description AS PackDesc,
  Item_1.Price AS PackPrice
FROM
  (
    (
      (
        Item
        LEFT JOIN UPCItem ON Item.SubPack = UPCItem.ID
      )
      LEFT JOIN Item AS Item_1 ON UPCItem.ItemID = Item_1.ID
    )
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  )
  LEFT JOIN SalesDepartment AS SalesDepartment_1 ON Item_1.SalesDepartmentID = SalesDepartment_1.ID
WHERE
  (
    (
      (Item.SalesDepartmentID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (Item.UnitsPerPack)> 1
    )
  )
ORDER BY
  Item.SalesDepartmentID,
  Item.Description;
