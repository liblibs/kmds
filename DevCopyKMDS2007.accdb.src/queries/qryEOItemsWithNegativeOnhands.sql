SELECT
  SalesGroup.ID AS [Group],
  SalesGroup.Description AS GroupDesc,
  SalesDepartment.ID AS Department,
  SalesDepartment.Description AS DepartmentDesc,
  Item.ID,
  Item.Description,
  qryUPCItemMax.UPCCodeMax AS UPCCode,
  Item.Onhand,
  Item.Price,
  SalesDepartment.NegativeOnhand,
  [Onhand] * [Price] AS Ext
FROM
  (
    SalesGroup
    INNER JOIN (
      SalesDepartment
      INNER JOIN Item ON SalesDepartment.ID = Item.SalesDepartmentID
    ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
  )
  LEFT JOIN qryUPCItemMax ON Item.ID = qryUPCItemMax.ItemID
WHERE
  (
    (
      (Item.Onhand)< 0
    )
    AND (
      (SalesDepartment.NegativeOnhand)<> 0
    )
  )
ORDER BY
  SalesGroup.ID,
  SalesDepartment.ID,
  Item.ID;
