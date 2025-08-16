SELECT
  SalesDepartment.ID AS SalesDptID,
  SalesDepartment.Description AS SalesDptDesc,
  IIf(
    IsNull(SubCategory.ID),
    "n/a",
    Subcategory.Description
  ) AS Category,
  Item.ID,
  Item.Description,
  Size.Description AS SizeDesc,
  Item.Onhand,
  Item.Price,
  qryItemsByDaysSupplySoldTotals.SumOfQuantity,
  qryItemsByDaysSupplySoldTotals.SumOfPrice,
  Item.SetupDate,
  Item.DateLastSold
FROM
  (
    (
      SalesDepartment
      INNER JOIN (
        Item
        INNER JOIN qryItemsByDaysSupplySoldTotals ON Item.ID = qryItemsByDaysSupplySoldTotals.Code
      ) ON SalesDepartment.ID = Item.SalesDepartmentID
    )
    INNER JOIN [Size] ON Item.SizeID = Size.ID
  )
  LEFT JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
WHERE
  (
    (
      (SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (
        IIf(
          IsNull(SubCategory.ID),
          "n/a",
          Subcategory.Description
        )
      ) Between Forms!PrintInventoryReport!cboCategoryIDBegin
      And Forms!PrintInventoryReport!cboCategoryIDEnd
    )
  )
ORDER BY
  SalesDepartment.ID;
