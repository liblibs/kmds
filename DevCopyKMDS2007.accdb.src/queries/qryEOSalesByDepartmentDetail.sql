SELECT
  IIf(
    IsNull(SalesGroup.id),
    "2",
    SalesGroup.GroupLevel
  ) AS GroupLvl,
  IIf(
    IsNull(SalesGroup.ID),
    999,
    SalesGroup.id
  ) AS [Group],
  IIf(
    IsNull(SalesGroup.ID),
    "Drops and Loans",
    SalesGroup.Description
  ) AS GroupDesc,
  IIf(
    IsNull(SalesDepartment.ID),
    0,
    SalesDepartment.ID
  ) AS Department,
  IIf(
    IsNull(SalesDepartment.ID),
    "N/A",
    IIf(
      Involine5.Transtype = 'C', SalesDepartment.Description & " Club Sales",
      SalesDepartment.Description
    )
  ) AS DepartmentDesc,
  Item.Description,
  Item.ID,
  Sum(InvoLine5.Quantity) AS TotalQ,
  Sum(InvoLine5.Price) AS TotalP,
  InvoLine5.TransType,
  Sum(InvoLine5.Cost) AS TotalC,
  Sum(InvoLine5.Price - InvoLine5.Cost) AS TotalGP,
  Sum(
    IIf([Quantity] < 0, [Quantity], 0)
  ) AS TotalReturnsQ,
  Sum(
    IIf([Quantity] < 0, InvoLine5.price, 0)
  ) AS TotalReturnsD,
  Sum(
    IIf([Quantity] > 0, [Quantity], 0)
  ) AS TotalSalesQ,
  Sum(
    IIf([Quantity] > 0, InvoLine5.price, 0)
  ) AS TotalSalesD
FROM
  SalesGroup
  RIGHT JOIN (
    SalesDepartment
    RIGHT JOIN (
      InvoLine5
      LEFT JOIN Item ON InvoLine5.Code = Item.ID
    ) ON SalesDepartment.ID = Item.SalesDepartmentID
  ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
WHERE
  (
    (
      (InvoLine5.TransType)<> 'X'
      And (InvoLine5.TransType)<> 'T'
    )
    AND (
      (InvoLine5.InputCode)<> "Cash"
      And (InvoLine5.InputCode)<> "Change"
      And (InvoLine5.InputCode)<> "Checks"
      And (InvoLine5.InputCode)<> "Visa MC"
    )
    AND (
      (InvoLine5.Price)<> 0
    )
  )
GROUP BY
  IIf(
    IsNull(SalesGroup.id),
    "2",
    SalesGroup.GroupLevel
  ),
  IIf(
    IsNull(SalesGroup.ID),
    999,
    SalesGroup.id
  ),
  IIf(
    IsNull(SalesGroup.ID),
    "Drops and Loans",
    SalesGroup.Description
  ),
  IIf(
    IsNull(SalesDepartment.ID),
    0,
    SalesDepartment.ID
  ),
  IIf(
    IsNull(SalesDepartment.ID),
    "N/A",
    IIf(
      Involine5.Transtype = 'C', SalesDepartment.Description & " Club Sales",
      SalesDepartment.Description
    )
  ),
  Item.Description,
  Item.ID,
  InvoLine5.TransType
HAVING
  (
    (
      (Item.ID) Is Not Null
    )
  )
ORDER BY
  IIf(
    IsNull(SalesGroup.ID),
    999,
    SalesGroup.id
  ),
  IIf(
    IsNull(SalesDepartment.ID),
    0,
    SalesDepartment.ID
  ),
  IIf(
    IsNull(SalesDepartment.ID),
    "N/A",
    IIf(
      Involine5.Transtype = 'C', SalesDepartment.Description & " Club Sales",
      SalesDepartment.Description
    )
  ),
  Item.ID;
