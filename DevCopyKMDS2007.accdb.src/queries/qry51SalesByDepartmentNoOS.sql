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
    Tender.ID,
    SalesDepartment.ID
  ) AS Department,
  IIf(
    IsNull(SalesDepartment.ID),
    Tender.Description,
    IIf(
      InvoLineAllRegisters.Transtype = 'C',
      SalesDepartment.Description & " Club Sales",
      SalesDepartment.Description
    )
  ) AS DepartmentDesc,
  Sum(
    IIf([TransType] = "T", 1, [Quantity])
  ) AS TotalQ,
  Sum(InvoLineAllRegisters.Price) AS TotalP,
  Sum(
    IIf(
      [CashsysID] = 1,
      IIf([TransType] = "T", 1, [Quantity]),
      0
    )
  ) AS Q1,
  Sum(
    IIf(
      [CashsysID] = 1, InvoLineAllRegisters.Price,
      0
    )
  ) AS D1,
  Sum(
    IIf(
      [CashsysID] = 2,
      IIf([TransType] = "T", 1, [Quantity]),
      0
    )
  ) AS Q2,
  Sum(
    IIf(
      [CashsysID] = 2, InvoLineAllRegisters.Price,
      0
    )
  ) AS D2,
  Sum(
    IIf(
      [CashsysID] = 3,
      IIf([TransType] = "T", 1, [Quantity]),
      0
    )
  ) AS Q3,
  Sum(
    IIf(
      [CashsysID] = 3, InvoLineAllRegisters.Price,
      0
    )
  ) AS D3,
  Sum(
    IIf(
      [CashsysID] = 4,
      IIf([TransType] = "T", 1, [Quantity]),
      0
    )
  ) AS Q4,
  Sum(
    IIf(
      [CashsysID] = 4, InvoLineAllRegisters.Price,
      0
    )
  ) AS D4,
  InvoLineAllRegisters.TransType,
  0 AS TotalC,
  Sum(InvoLineAllRegisters.Price) AS TotalGP,
  Sum(
    IIf(
      [CashsysID] = 5,
      IIf([TransType] = "T", 1, [Quantity]),
      0
    )
  ) AS Q5,
  Sum(
    IIf(
      [CashsysID] = 5, [InvoLineAllRegisters].[Price],
      0
    )
  ) AS D5
FROM
  (
    SalesGroup
    RIGHT JOIN (
      SalesDepartment
      RIGHT JOIN (
        InvoLineAllRegisters
        LEFT JOIN Item ON InvoLineAllRegisters.Code = Item.ID
      ) ON SalesDepartment.ID = Item.SalesDepartmentID
    ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
  )
  LEFT JOIN Tender ON InvoLineAllRegisters.Code = Tender.ID
WHERE
  (
    (
      (InvoLineAllRegisters.TransType)<> 'X'
    )
    AND (
      (InvoLineAllRegisters.InputCode)<> "Cash"
      And (InvoLineAllRegisters.InputCode)<> "Change"
      And (InvoLineAllRegisters.InputCode)<> "Checks"
      And (InvoLineAllRegisters.InputCode)<> "Visa MC"
      And (InvoLineAllRegisters.InputCode) Not Like "Auth:*"
    )
    AND (
      (InvoLineAllRegisters.Price)<> 0
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
    Tender.ID,
    SalesDepartment.ID
  ),
  IIf(
    IsNull(SalesDepartment.ID),
    Tender.Description,
    IIf(
      InvoLineAllRegisters.Transtype = 'C',
      SalesDepartment.Description & " Club Sales",
      SalesDepartment.Description
    )
  ),
  InvoLineAllRegisters.TransType
ORDER BY
  IIf(
    IsNull(SalesGroup.ID),
    999,
    SalesGroup.id
  ),
  IIf(
    IsNull(SalesDepartment.ID),
    Tender.ID,
    SalesDepartment.ID
  ),
  IIf(
    IsNull(SalesDepartment.ID),
    Tender.Description,
    IIf(
      InvoLineAllRegisters.Transtype = 'C',
      SalesDepartment.Description & " Club Sales",
      SalesDepartment.Description
    )
  );
