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
  Item.DateLastSold,
  IIf(
    Forms!PrintInventoryReport!chkPrintZeroOnhands,
    -1, 0
  ) AS Expr1
FROM
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
WHERE
  (
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
      (Item.DateLastSold)< Forms!PrintInventoryReport!txtDateRangeFrom
    )
  )
ORDER BY
  SalesGroup.ID,
  SalesDepartment.ID,
  Item.ID;
