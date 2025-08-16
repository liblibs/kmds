SELECT
  SalesGroup.GroupLevel,
  SalesGroup.ID AS [Group],
  SalesGroup.Description AS GroupDesc,
  SalesDepartment.ID AS Department,
  IIf(
    involine.Transtype = 'C', "Club Sales",
    SalesDepartment.Description
  ) AS DepartmentDesc,
  Sum(InvoLine.Quantity) AS TotalQ,
  Sum(InvoLine.Price) AS Sales,
  Sum(
    IIf([CashsysID] = 1, [Quantity], 0)
  ) AS Q1,
  Sum(
    IIf([CashsysID] = 1, involine.Price, 0)
  ) AS D1,
  Sum(
    IIf([CashsysID] = 2, [Quantity], 0)
  ) AS Q2,
  Sum(
    IIf([CashsysID] = 2, involine.Price, 0)
  ) AS D2,
  Sum(
    IIf([CashsysID] = 3, [Quantity], 0)
  ) AS Q3,
  Sum(
    IIf([CashsysID] = 3, involine.Price, 0)
  ) AS D3,
  Sum(
    IIf([CashsysID] = 4, [Quantity], 0)
  ) AS Q4,
  Sum(
    IIf([CashsysID] = 4, involine.Price, 0)
  ) AS D4,
  InvoLine.TransType
FROM
  InvoLine
  INNER JOIN (
    SalesGroup
    INNER JOIN (
      SalesDepartment
      INNER JOIN Item ON SalesDepartment.ID = Item.SalesDepartmentID
    ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
  ) ON InvoLine.Code = Item.ID
WHERE
  (
    (
      (InvoLine.SetupDate) Between [Forms]![PrintInventoryReport]![txtDateRangeFrom]
      And [Forms]![PrintInventoryReport]![txtDateRangeTo]
    )
  )
GROUP BY
  SalesGroup.GroupLevel,
  SalesGroup.ID,
  SalesGroup.Description,
  SalesDepartment.ID,
  IIf(
    involine.Transtype = 'C', "Club Sales",
    SalesDepartment.Description
  ),
  InvoLine.TransType
HAVING
  (
    (
      (InvoLine.TransType)= "S"
      Or (InvoLine.TransType)= "C"
      Or (InvoLine.TransType)= "D"
    )
  )
ORDER BY
  SalesGroup.ID,
  SalesDepartment.ID;
