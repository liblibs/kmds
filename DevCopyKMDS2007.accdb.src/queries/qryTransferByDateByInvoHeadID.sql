SELECT
  First(InvoLine.SetupDate) AS FirstOfSetupDate,
  InvoLine.CashSysID,
  InvoLine.InvoHeadID,
  InvoLine.TransType,
  InvoLine.ShiftID,
  InvoLine.InputCode,
  Sum(
    IIf(
      Item.SalesDepartmentID = 1, [Quantity],
      0
    )
  ) AS LiquorQuantity,
  Sum(
    IIf(
      Item.SalesDepartmentID = 1, Involine.price,
      0
    )
  ) AS LiquorRetail,
  Sum(
    IIf(
      Item.SalesDepartmentID = 1, Involine.Cost,
      0
    )
  ) AS LiquorCost,
  Sum(
    IIf(
      Item.SalesDepartmentID <> 1, [Quantity],
      0
    )
  ) AS NonLiquorQuantity,
  Sum(
    IIf(
      Item.SalesDepartmentID <> 1, Involine.price,
      0
    )
  ) AS NonLiquorRetail,
  Sum(
    IIf(
      Item.SalesDepartmentID <> 1, Involine.Cost,
      0
    )
  ) AS NonLiquorCost,
  IIf(
    [LiquorRetail] < 0
    Or [NonLiquorRetail] < 0,
    "Out", "In"
  ) AS InOrOut,
  Sum(InvoLine.Quantity) AS TotalQuanity,
  Sum(InvoLine.Price) AS TotalRetail,
  Sum(InvoLine.Cost) AS TotalCost
FROM
  InvoLine
  INNER JOIN Item ON InvoLine.Code = Item.ID
WHERE
  (
    (
      (Item.SalesDepartmentID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
      Or (Item.SalesDepartmentID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (InvoLine.Drawer)= 0
    )
  )
GROUP BY
  InvoLine.CashSysID,
  InvoLine.InvoHeadID,
  InvoLine.TransType,
  InvoLine.ShiftID,
  InvoLine.InputCode
HAVING
  (
    (
      (
        First(InvoLine.SetupDate)
      ) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
    And (
      (InvoLine.TransType)= "X"
    )
  )
ORDER BY
  First(InvoLine.SetupDate);
