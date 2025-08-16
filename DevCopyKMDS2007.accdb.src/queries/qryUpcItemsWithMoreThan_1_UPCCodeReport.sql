SELECT
  SalesDepartment.ID AS SalesDpt,
  SalesDepartment.Description AS SalesDptDesc,
  SubCategory.Description AS Category,
  qryUpcItemsWithMoreThan_1_UPCCode.ItemID,
  Item.Description,
  Item.Onhand,
  Item.Price,
  UPCItem.ID AS UPCCode,
  UPCItem.SetupDate
FROM
  (
    (
      (
        qryUpcItemsWithMoreThan_1_UPCCode
        INNER JOIN UPCItem ON qryUpcItemsWithMoreThan_1_UPCCode.ItemID = UPCItem.ItemID
      )
      INNER JOIN Item ON UPCItem.ItemID = Item.ID
    )
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  )
  LEFT JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
WHERE
  (
    (
      (SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (UPCItem.ID)<> CStr(item.id)
    )
  )
ORDER BY
  SalesDepartment.ID,
  SubCategory.Description,
  Item.Description;
