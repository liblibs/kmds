SELECT
  InvoLine.SetupDate,
  InvoLine.CashSysID,
  InvoLine.InvoHeadID,
  InvoLine.TransType,
  InvoLine.ShiftID,
  InvoLine.InputCode,
  IIf(
    Item.SalesDepartmentID = 1, [Quantity],
    0
  ) AS LiquorQuantity,
  IIf(
    Item.SalesDepartmentID = 1, Involine.price,
    0
  ) AS LiquorRetail,
  IIf(
    Item.SalesDepartmentID = 1, Involine.Cost,
    0
  ) AS LiquorCost,
  IIf(
    Item.SalesDepartmentID <> 1, [Quantity],
    0
  ) AS NonLiquorQuantity,
  IIf(
    Item.SalesDepartmentID <> 1, Involine.price,
    0
  ) AS NonLiquorRetail,
  IIf(
    Item.SalesDepartmentID <> 1, Involine.Cost,
    0
  ) AS NonLiquorCost,
  IIf(
    [LiquorRetail] < 0
    Or [NonLiquorRetail] < 0,
    "Out", "In"
  ) AS InOrOut,
  InvoLine.Quantity AS TotalQuanity,
  InvoLine.Price AS TotalRetail,
  InvoLine.Cost AS TotalCost,
  Item.ID,
  Item.Description
FROM
  InvoLine
  INNER JOIN Item ON InvoLine.Code = Item.ID
WHERE
  (
    (
      (InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
    And (
      (InvoLine.TransType)= "X"
    )
    And (
      (Item.SalesDepartmentID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
      Or (Item.SalesDepartmentID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (InvoLine.Drawer)= 0
    )
  )
ORDER BY
  InvoLine.SetupDate;
