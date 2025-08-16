SELECT
  SalesGroup.ID AS [Group],
  SalesGroup.Description AS GroupDesc,
  SalesDepartment.ID AS Department,
  SalesDepartment.Description AS DepartmentDesc,
  Item.ID,
  Item.Description,
  UPCItem.ID AS UPCCode,
  Item.Price,
  Item.SetupDate,
  Format(
    Now()-1,
    "mm/dd/yyyy"
  ) AS SetupSince
FROM
  (
    SalesGroup
    INNER JOIN (
      SalesDepartment
      INNER JOIN Item ON SalesDepartment.ID = Item.SalesDepartmentID
    ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
  )
  LEFT JOIN UPCItem ON Item.ID = UPCItem.ItemID
WHERE
  (
    (
      (Item.SetupDate)> Format(
        Now()-1,
        "mm/dd/yyyy"" 23"":""59"":""59"""
      )
    )
  )
ORDER BY
  SalesGroup.ID,
  SalesDepartment.ID,
  Item.ID;
