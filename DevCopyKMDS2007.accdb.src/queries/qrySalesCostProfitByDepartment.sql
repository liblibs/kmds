SELECT
  SalesGroup.GroupLevel,
  SalesGroup.ID AS [Group],
  SalesGroup.Description AS GroupDesc,
  SalesDepartment.ID AS Department,
  IIf(
    involine.Transtype = 'C',
    Trim(SalesDepartment.Description)& "-Club Sales",
    SalesDepartment.Description
  ) AS DepartmentDesc,
  Sum(InvoLine.Quantity) AS TotalQ,
  Sum(InvoLine.Price) AS TotalP,
  Sum(
    (InvoLine.Cost)
  ) AS TotalC,
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
    involine.Transtype = 'C',
    Trim(SalesDepartment.Description)& "-Club Sales",
    SalesDepartment.Description
  ),
  InvoLine.TransType
HAVING
  (
    (
      (SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (InvoLine.TransType)= "S"
      Or (InvoLine.TransType)= "C"
      Or (InvoLine.TransType)= "D"
    )
  )
ORDER BY
  SalesGroup.ID,
  SalesDepartment.ID,
  IIf(
    involine.Transtype = 'C',
    Trim(SalesDepartment.Description)& "-Club Sales",
    SalesDepartment.Description
  );
