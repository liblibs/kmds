SELECT
  SalesGroup.ID AS [Group],
  SalesGroup.Description AS GroupDesc,
  SalesDepartment.ID AS Department,
  SalesDepartment.Description AS DepartmentDesc,
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
  Item.TPR
FROM
  SalesGroup
  INNER JOIN (
    SalesDepartment
    INNER JOIN (
      Item
      LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID
    ) ON SalesDepartment.ID = Item.SalesDepartmentID
  ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
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
      (Item.TPR)<> 0
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
      (Item.TPR)<> 0
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
