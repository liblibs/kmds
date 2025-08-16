SELECT
  SalesGroup.ID AS [Group],
  SalesGroup.Description AS GroupDesc,
  SalesDepartment.ID AS Department,
  SalesDepartment.Description AS DepartmentDesc,
  Category.ID AS CategoryID,
  Category.Description AS CategoryDescription,
  SubCategory.Description AS SubCategoryDescription,
  UPCItem.ID AS AlphaUPC,
  Item.ID,
  Item.Description,
  Item.Onhand,
  Item.Price,
  Item.Cost,
  SalesDepartment.IncludeInInventoryValueReports,
  SalesDepartment.TargetGP,
  IIf(
    Forms!PrintInventoryReport!chkPrintZeroOnhands,
    -1, 0
  ) AS PrintZeroOnhands,
  Item.DateLastSold,
  Item.StockNumber
FROM
  (
    SalesGroup
    INNER JOIN (
      SalesDepartment
      INNER JOIN (
        Category
        RIGHT JOIN (
          Item
          LEFT JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
        ) ON Category.ID = SubCategory.CategoryID
      ) ON SalesDepartment.ID = Item.SalesDepartmentID
    ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
  )
  INNER JOIN UPCItem ON Item.ID = UPCItem.ItemID
WHERE
  (
    (
      (SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (UPCItem.ID)> "A"
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
      (UPCItem.ID)> "A"
    )
    And (
      (
        SalesDepartment.IncludeInInventoryValueReports
      )<> 0
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
  Category.ID,
  SubCategory.Description,
  UPCItem.ID;
