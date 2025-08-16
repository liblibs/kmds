SELECT
  Format(
    InvoLine.SetupDate, "mm/dd/yyyy"
  ) AS SD,
  SalesDepartment.SalesGroupID,
  SalesDepartment.ID,
  SalesDepartment.Description,
  Item.ID,
  Item.Description,
  InvoLine.Quantity,
  IIf(
    [Quantity] = 0, 0, InvoLine.Price / [Quantity]
  ) AS UPrice,
  InvoLine.Price AS Ext,
  Cashier.ID,
  Cashier.Name,
  InvoLine.SetupDate,
  InvoLine.TransType,
  InvoLine.InputCode,
  InvoLine.SalesGroup AS NewOnShelf
FROM
  (
    InvoLine
    INNER JOIN Cashier ON InvoLine.CashierID = Cashier.ID
  )
  INNER JOIN (
    Item
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  ) ON InvoLine.Code = Item.ID
WHERE
  (
    (
      (SalesDepartment.ID) Between Forms!PrintInventoryAuditReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryAuditReport!cboSalesDepartmentIDEnd
    )
    And (
      (Cashier.ID) Between Forms!PrintInventoryAuditReport!cboCashiersIDBegin
      And Forms!PrintInventoryAuditReport!cboCashiersIDEnd
    )
    And (
      (InvoLine.SetupDate) Between Forms!PrintInventoryAuditReport!dtpDateBegin
      And Forms!PrintInventoryAuditReport!dtpDateEnd
    )
    And (
      (InvoLine.TransType)= "A"
      Or (InvoLine.TransType)= "R"
    )
  )
ORDER BY
  SalesDepartment.SalesGroupID,
  SalesDepartment.ID,
  Item.ID,
  Cashier.ID;
