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
  InvoLine.TransType
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
      (InvoLine.TransType)= "A"
    )
    And (
      (
        Format(
          InvoLine.SetupDate, "mm/dd/yyyy"
        )
      )= Format(
        Now(),
        "mm/dd/yyyy"
      )
    )
  )
ORDER BY
  SalesDepartment.SalesGroupID,
  SalesDepartment.ID,
  Item.ID,
  Cashier.ID;
