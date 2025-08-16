SELECT
  SalesGroup.ID AS [Group],
  SalesGroup.Description AS GroupDesc,
  SalesDepartment.ID AS Department,
  SalesDepartment.Description AS DepartmentDesc,
  Category.ID AS CategoryID,
  Category.Description AS CategoryDescription,
  SubCategory.Description AS SubCategoryDescription,
  Item.ID,
  Item.Description,
  Item.Onhand,
  Item.Price,
  qryItemsSalePriceNow.SalePrice,
  Item.Cost,
  SalesDepartment.IncludeInInventoryValueReports,
  SalesDepartment.TargetGP,
  IIf(
    Forms!PrintInventoryReport!chkPrintZeroOnhands,
    -1, 0
  ) AS PrintZeroOnhands,
  Size.Oz,
  IIf([OZ] > 0, [Item].[Price] / [OZ], 0) AS PPOz,
  [onhand] * [Item].[Price] AS ExtPrice,
  dbo_OLCCitem.StockStatus
FROM
  (
    (
      SalesGroup
      INNER JOIN (
        SalesDepartment
        INNER JOIN (
          Category
          RIGHT JOIN (
            (
              Item
              LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID
            )
            LEFT JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
          ) ON Category.ID = SubCategory.CategoryID
        ) ON SalesDepartment.ID = Item.SalesDepartmentID
      ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
    )
    LEFT JOIN [Size] ON Item.SizeID = Size.ID
  )
  INNER JOIN dbo_OLCCitem ON Item.ID = dbo_OLCCitem.Code
WHERE
  (
    (
      (SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (Item.Onhand)<> 0
    )
    And (
      (
        SalesDepartment.IncludeInInventoryValueReports
      )<> 0
    )
    And (
      (dbo_OLCCitem.StockStatus)= "S"
    )
    And (
      (
        IIf(
          Forms!PrintInventoryReport!chkPrintZeroOnhands,
          -1, 0
        )
      )= 0
    )
  )
  Or (
    (
      (SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (
        SalesDepartment.IncludeInInventoryValueReports
      )<> 0
    )
    And (
      (dbo_OLCCitem.StockStatus)= "S"
    )
    And (
      (
        IIf(
          Forms!PrintInventoryReport!chkPrintZeroOnhands,
          -1, 0
        )
      )=-1
    )
  )
ORDER BY
  SalesGroup.ID,
  SalesDepartment.ID,
  Item.ID;
