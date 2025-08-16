SELECT
  Category.SalesDepartmentID,
  SalesDepartment.Description,
  Category.Description AS CatDescription,
  SubCategory.Description AS SubCatDesc,
  SubCategory.TargetGP,
  IIf(
    IsNull(
      [qryItemsByCategoryCounts].[count]
    ),
    0,
    [qryItemsByCategoryCounts].[count]
  ) AS Items
FROM
  (
    (
      SubCategory
      INNER JOIN Category ON SubCategory.CategoryID = Category.ID
    )
    INNER JOIN SalesDepartment ON Category.SalesDepartmentID = SalesDepartment.ID
  )
  LEFT JOIN qryItemsByCategoryCounts ON SubCategory.ID = qryItemsByCategoryCounts.SubCategoryID
ORDER BY
  Category.SalesDepartmentID,
  Category.Description,
  SubCategory.Description;
