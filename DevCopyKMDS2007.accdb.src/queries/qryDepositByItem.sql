SELECT
  Item.SalesDepartmentID,
  SalesDepartment.Description,
  Item.ID,
  Item.Description,
  Item.Onhand,
  Item.Deposit,
  [onhand] * [deposit] AS DepositTotal,
  SubCategory.Description AS SubCategoryDescription
FROM
  (
    SalesDepartment
    INNER JOIN Item ON SalesDepartment.ID = Item.SalesDepartmentID
  )
  INNER JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
WHERE
  (
    (
      (Item.Onhand)<> 0
    )
    AND (
      (Item.Deposit)<> 0
    )
  )
ORDER BY
  Item.SalesDepartmentID,
  Item.Description;
