DELETE Category.SalesDepartmentID,
SubCategory.*,
SubCategory.ID
FROM
  SubCategory
  INNER JOIN Category ON SubCategory.CategoryID = Category.ID
WHERE
  (
    (
      (Category.SalesDepartmentID)= 1
    )
    AND (
      (SubCategory.ID)> 0
    )
  );
