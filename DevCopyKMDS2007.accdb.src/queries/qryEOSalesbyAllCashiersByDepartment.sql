SELECT
  SalesGroup.GroupLevel,
  InvoLine5.CashierID,
  Cashier.Name,
  IIf(
    [TransType] = "T", "Loans and Drops",
    " Sales"
  ) AS SalesOrTender,
  IIf(
    IsNull(SalesGroup.ID),
    999,
    SalesGroup.ID
  ) AS [Group],
  IIf(
    IsNull(SalesGroup.Description),
    "Loans and Drops",
    SalesGroup.Description
  ) AS GroupDesc,
  IIf(
    IsNull(SalesDepartment.ID),
    Tender.ID,
    SalesDepartment.ID
  ) AS Department,
  IIf(
    IsNull(SalesDepartment.Description),
    Tender.Description,
    IIf(
      Involine5.Transtype = 'C', SalesDepartment.Description & " Club Sales",
      SalesDepartment.Description
    )
  ) AS DepartmentDesc,
  Sum(
    IIf([TransType] = "T", 0, [Quantity])
  ) AS TotalQ,
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
  Sum(
    IIf([CashsysID] = 5, [Quantity], 0)
  ) AS Q5,
  Sum(
    IIf(
      [CashsysID] = 5, [InvoLine5].[Price],
      0
    )
  ) AS D5
FROM
  (
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
    INNER JOIN Cashier ON InvoLine5.CashierID = Cashier.ID
  )
  LEFT JOIN Tender ON InvoLine5.Code = Tender.ID
WHERE
  (
    (
      (InvoLine5.InputCode)<> "Cash"
      And (InvoLine5.InputCode)<> "Change"
      And (InvoLine5.InputCode)<> "Checks"
      And (InvoLine5.InputCode)<> "Visa MC"
      And (InvoLine5.InputCode) Not Like "Auth:*"
    )
    AND (
      (InvoLine5.Price)<> 0
    )
  )
GROUP BY
  SalesGroup.GroupLevel,
  InvoLine5.CashierID,
  Cashier.Name,
  IIf(
    [TransType] = "T", "Loans and Drops",
    " Sales"
  ),
  IIf(
    IsNull(SalesGroup.ID),
    999,
    SalesGroup.ID
  ),
  IIf(
    IsNull(SalesGroup.Description),
    "Loans and Drops",
    SalesGroup.Description
  ),
  IIf(
    IsNull(SalesDepartment.ID),
    Tender.ID,
    SalesDepartment.ID
  ),
  IIf(
    IsNull(SalesDepartment.Description),
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
      (InvoLine5.TransType)<> "X"
    )
  )
ORDER BY
  InvoLine5.CashierID,
  IIf(
    IsNull(SalesGroup.ID),
    999,
    SalesGroup.ID
  ),
  IIf(
    IsNull(SalesDepartment.ID),
    Tender.ID,
    SalesDepartment.ID
  );
