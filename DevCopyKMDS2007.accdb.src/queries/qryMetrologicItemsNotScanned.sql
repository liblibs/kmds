SELECT
  Item.SalesDepartmentID,
  SalesDepartment.Description,
  SubCategory.ID,
  SubCategory.Description,
  Item.ID AS ItemID,
  Item.Description AS ItemDescription,
  Item.Onhand,
  qryMetrologicScanBatchTotalByItemID.SumOfQuantity,
  Item.Price
FROM
  (
    (
      Item
      INNER JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
    )
    LEFT JOIN qryMetrologicScanBatchTotalByItemID ON Item.ID = qryMetrologicScanBatchTotalByItemID.ItemID
  )
  INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
WHERE
  (
    (
      (Item.SalesDepartmentID) Between Forms!MetrologicOptimiser!cboSalesDepartmentIDBegin
      And Forms!MetrologicOptimiser!cboSalesdepartmentIDEnd
    )
    And (
      (SubCategory.Description) Between Forms!MetrologicOptimiser!cboNotScannedCategoryIDBegin
      And Forms!MetrologicOptimiser!cboNotScannedCategoryIDEnd
    )
    And (
      (Item.Onhand)<> 0
    )
    And (
      (
        qryMetrologicScanBatchTotalByItemID.SumOfQuantity
      ) Is Null
    )
  )
ORDER BY
  Item.SalesDepartmentID,
  SubCategory.Description,
  Item.Description;
