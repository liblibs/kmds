SELECT
  InvoLine.CashierID,
  SalesDepartment.ID AS SalesDptID,
  InvoLine.SetupDate,
  Cashier.Name AS CashierName,
  SalesDepartment.Description AS SalesDptDesc,
  IIf(
    IsNull(SubCategory.ID),
    "n/a",
    Subcategory.Description
  ) AS Category,
  Item.ID,
  Item.Description,
  Size.Description AS SizeDesc,
  Item.Onhand,
  InvoLine.Quantity,
  InvoLine.Price,
  InvoLine.CashSysID AS Register,
  InvoLine.InvoHeadID AS Invoice
FROM
  (
    (
      (
        (
          SalesDepartment
          INNER JOIN Item ON SalesDepartment.ID = Item.SalesDepartmentID
        )
        INNER JOIN [Size] ON Item.SizeID = Size.ID
      )
      LEFT JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
    )
    INNER JOIN InvoLine ON Item.ID = InvoLine.Code
  )
  INNER JOIN Cashier ON InvoLine.CashierID = Cashier.ID
WHERE
  (
    (
      (SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
  )
ORDER BY
  InvoLine.CashierID,
  SalesDepartment.ID,
  InvoLine.SetupDate;
