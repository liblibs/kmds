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
      Involine5.Transtype = 'C', SalesDepartment.Description & " Club Sales",
      SalesDepartment.Description
    )
  ) AS DepartmentDesc,
  Sum(InvoLine5.Quantity) AS TotalQ,
  Sum(InvoLine5.Price) AS TotalP,
  Sum(
    IIf([CashsysID] = 1, [Quantity], 0)
  ) AS Q1,
  Sum(
    IIf(
      [CashsysID] = 1, InvoLine5.Price, 0
    )
  ) AS D1,
  Sum(
    IIf([CashsysID] = 2, [Quantity], 0)
  ) AS Q2,
  Sum(
    IIf(
      [CashsysID] = 2, InvoLine5.Price, 0
    )
  ) AS D2,
  Sum(
    IIf([CashsysID] = 3, [Quantity], 0)
  ) AS Q3,
  Sum(
    IIf(
      [CashsysID] = 3, InvoLine5.Price, 0
    )
  ) AS D3,
  Sum(
    IIf([CashsysID] = 4, [Quantity], 0)
  ) AS Q4,
  Sum(
    IIf(
      [CashsysID] = 4, InvoLine5.Price, 0
    )
  ) AS D4,
  InvoLine5.TransType,
  Sum(InvoLine5.Cost) AS TotalC,
  Sum(InvoLine5.Price - InvoLine5.Cost) AS TotalGP
FROM
  (
    SalesGroup
    RIGHT JOIN (
      SalesDepartment
      RIGHT JOIN (
        InvoLine5
        LEFT JOIN Item ON InvoLine5.Code = Item.ID
      ) ON SalesDepartment.ID = Item.SalesDepartmentID
    ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
  )
  LEFT JOIN Tender ON InvoLine5.Code = Tender.ID
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
    Tender.ID,
    SalesDepartment.ID
  ),
  IIf(
    IsNull(SalesDepartment.ID),
    Tender.Description,
    IIf(
      Involine5.Transtype = 'C', SalesDepartment.Description & " Club Sales",
      SalesDepartment.Description
    )
  ),
  InvoLine5.TransType
HAVING
  (
    (
      (
        IIf(
          IsNull([SalesDepartment].[ID]),
          [Tender].[Description],
          IIf(
            [Involine5].[Transtype] = 'C', [SalesDepartment].[Description] & " Club Sales",
            [SalesDepartment].[Description]
          )
        )
      )<> "Visa MC"
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
    Tender.ID,
    SalesDepartment.ID
  ),
  IIf(
    IsNull(SalesDepartment.ID),
    Tender.Description,
    IIf(
      Involine5.Transtype = 'C', SalesDepartment.Description & " Club Sales",
      SalesDepartment.Description
    )
  );
