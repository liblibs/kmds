SELECT
  SalesGroup.ID AS [Group],
  SalesGroup.Description AS GroupDesc,
  SalesDepartment.ID AS Department,
  SalesDepartment.Description AS DepartmentDesc,
  Item.ID,
  Item.Description,
  Sum(InvoLine5.Quantity) AS SoldToday,
  Sum(InvoLine5.Price) AS DollarsSoldToday,
  InvoLine5.TransType
FROM
  InvoLine5
  INNER JOIN (
    SalesGroup
    INNER JOIN (
      SalesDepartment
      INNER JOIN Item ON SalesDepartment.ID = Item.SalesDepartmentID
    ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
  ) ON InvoLine5.Code = Item.ID
GROUP BY
  SalesGroup.ID,
  SalesGroup.Description,
  SalesDepartment.ID,
  SalesDepartment.Description,
  Item.ID,
  Item.Description,
  InvoLine5.TransType
HAVING
  (
    (
      (InvoLine5.TransType)= "S"
      Or (InvoLine5.TransType)= "D"
    )
  )
ORDER BY
  SalesGroup.ID,
  SalesDepartment.ID,
  Item.ID;
