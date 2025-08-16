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
  Size.Oz,
  IIf([OZ] > 0, [Price] / [OZ], 0) AS PPOz
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
  LEFT JOIN [Size] ON Item.SizeID = Size.ID
ORDER BY
  SalesGroup.ID,
  SalesDepartment.ID,
  Item.ID;
