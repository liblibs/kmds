SELECT
  SalesGroup.ID AS [Group],
  SalesGroup.Description AS GroupDesc,
  SalesDepartment.ID AS Department,
  SalesDepartment.Description AS DepartmentDesc,
  Item.ID,
  Item.Description,
  Item.Onhand,
  Item.Price,
  Item.AvgTurns,
  Item.UnitsPerOrder,
  Item.UnitsPerCase
FROM
  SalesGroup
  INNER JOIN (
    SalesDepartment
    INNER JOIN Item ON SalesDepartment.ID = Item.SalesDepartmentID
  ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
WHERE
  (
    (
      (Item.UnitsPerOrder)<> [AvgTurns]
    )
  )
ORDER BY
  SalesGroup.ID,
  SalesDepartment.ID,
  Item.ID;
