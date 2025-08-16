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
  ItemMovement.CurrMM,
  ItemMovement.Jan,
  ItemMovement.Feb,
  ItemMovement.Mar,
  ItemMovement.Apr,
  ItemMovement.May,
  ItemMovement.Jun,
  ItemMovement.Jul,
  ItemMovement.Aug,
  ItemMovement.Sep,
  ItemMovement.Oct,
  ItemMovement.Nov,
  ItemMovement.Dec,
  ItemMovement.YTDQty AS Total
FROM
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
  INNER JOIN ItemMovement ON Item.ID = ItemMovement.Code
WHERE
  (
    (
      (SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (SubCategory.Description) Between Forms!PrintInventoryReport!cboCategoryIDBegin
      And Forms!PrintInventoryReport!cboCategoryIDEnd
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
      (ItemMovement.Description)= "  Sales -- Total"
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
      (SubCategory.Description) Between Forms!PrintInventoryReport!cboCategoryIDBegin
      And Forms!PrintInventoryReport!cboCategoryIDEnd
    )
    And (
      (
        SalesDepartment.IncludeInInventoryValueReports
      )<> 0
    )
    And (
      (ItemMovement.Description)= "  Sales -- Total"
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
  Item.ID;
