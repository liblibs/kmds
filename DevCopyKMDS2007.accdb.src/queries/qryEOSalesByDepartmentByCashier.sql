PARAMETERS parCashier Long;
SELECT
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
    IIf(
      InvoLine5.CashsysID = 1, [Quantity],
      0
    )
  ) AS Q1,
  Sum(
    IIf(
      InvoLine5.CashsysID = 1, InvoLine5.Price,
      0
    )
  ) AS D1,
  Sum(
    IIf(
      InvoLine5.CashsysID = 2, [Quantity],
      0
    )
  ) AS Q2,
  Sum(
    IIf(
      InvoLine5.CashsysID = 2, InvoLine5.Price,
      0
    )
  ) AS D2,
  Sum(
    IIf(
      InvoLine5.CashsysID = 3, [Quantity],
      0
    )
  ) AS Q3,
  Sum(
    IIf(
      InvoLine5.CashsysID = 3, InvoLine5.Price,
      0
    )
  ) AS D3,
  Sum(
    IIf(
      InvoLine5.CashsysID = 4, [Quantity],
      0
    )
  ) AS Q4,
  Sum(
    IIf(
      InvoLine5.CashsysID = 4, InvoLine5.Price,
      0
    )
  ) AS D4,
  InvoLine5.TransType,
  InvoLine5.CashierID
FROM
  SalesGroup
  INNER JOIN (
    SalesDepartment
    INNER JOIN (
      InvoLine5
      INNER JOIN Item ON InvoLine5.Code = Item.ID
    ) ON SalesDepartment.ID = Item.SalesDepartmentID
  ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
WHERE
  (
    (
      (InvoLine5.TransType)<> 'T'
    )
  )
GROUP BY
  SalesGroup.ID,
  SalesGroup.Description,
  SalesDepartment.ID,
  IIf(
    Involine5.Transtype = 'C', "Club Sales",
    SalesDepartment.Description
  ),
  InvoLine5.TransType,
  InvoLine5.CashierID
HAVING
  (
    (
      (InvoLine5.CashierID)= [parCashier]
    )
  )
ORDER BY
  SalesGroup.ID,
  SalesDepartment.ID;
