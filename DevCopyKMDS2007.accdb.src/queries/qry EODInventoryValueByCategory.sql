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
  SalesDepartment.TargetGP
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
      (
        SalesDepartment.IncludeInInventoryValueReports
      )<> 0
    )
  )
ORDER BY
  SalesGroup.ID,
  SalesDepartment.ID,
  Category.ID,
  SubCategory.Description,
  Item.ID;
