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
  SalesDepartment.Description AS DepartmentDesc,
  Item.Description,
  Item.ID,
  Sum(Involine.Quantity) AS TotalQ,
  Sum(Involine.Price) AS TotalP,
  Sum(Involine.Cost) AS TotalC,
  Sum(Involine.Price - Involine.Cost) AS TotalGP,
  Sum(
    IIf([Quantity] < 0, [Quantity], 0)
  ) AS TotalReturnsQ,
  Sum(
    IIf([Quantity] < 0, Involine.price, 0)
  ) AS TotalReturnsD,
  Sum(
    IIf([Quantity] > 0, [Quantity], 0)
  ) AS TotalSalesQ,
  Sum(
    IIf([Quantity] > 0, Involine.price, 0)
  ) AS TotalSalesD
FROM
  SalesGroup
  INNER JOIN (
    SalesDepartment
    INNER JOIN (
      Involine
      INNER JOIN Item ON Involine.Code = Item.ID
    ) ON SalesDepartment.ID = Item.SalesDepartmentID
  ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
WHERE
  (
    (
      (SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (Involine.TransType)= "S"
      Or (Involine.TransType)= "C"
    )
    And (
      (Involine.Price)<> 0
    )
    And (
      (Involine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
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
  SalesDepartment.Description,
  Item.Description,
  Item.ID
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
  SalesDepartment.Description,
  Item.ID;
