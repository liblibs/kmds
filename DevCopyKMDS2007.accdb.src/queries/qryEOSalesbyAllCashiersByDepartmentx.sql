SELECT
  InvoLine5.CashierID,
  Cashier.Name,
  SalesGroup.ID AS [Group],
  SalesGroup.Description AS GroupDesc,
  SalesDepartment.ID AS Department,
  IIf(
    Involine5.Transtype = 'C', "Club Sales",
    SalesDepartment.Description
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
  InvoLine5.TransType
FROM
  (
    SalesGroup
    INNER JOIN (
      SalesDepartment
      INNER JOIN (
        InvoLine5
        INNER JOIN Item ON InvoLine5.Code = Item.ID
      ) ON SalesDepartment.ID = Item.SalesDepartmentID
    ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
  )
  INNER JOIN Cashier ON InvoLine5.CashierID = Cashier.ID
WHERE
  (
    (
      (InvoLine5.TransType)<> 'T'
    )
  )
GROUP BY
  InvoLine5.CashierID,
  Cashier.Name,
  SalesGroup.ID,
  SalesGroup.Description,
  SalesDepartment.ID,
  IIf(
    Involine5.Transtype = 'C', "Club Sales",
    SalesDepartment.Description
  ),
  InvoLine5.TransType
ORDER BY
  InvoLine5.CashierID,
  SalesGroup.ID,
  SalesDepartment.ID;
