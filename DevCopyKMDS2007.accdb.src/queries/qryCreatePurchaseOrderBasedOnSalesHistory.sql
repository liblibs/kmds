PARAMETERS parBeginDate DateTime,
parEndDate DateTime,
parVendorID Long,
parBeginSalesDepartmentID Long,
parEndSalesDepartmentID Long;
SELECT
  Item.VendorID,
  Item.ID AS ItemID,
  SUM(InvoLine.Quantity) AS QuantityOrdered,
  SUM(InvoLine.Quantity) AS QuantityShipped,
  IIf(
    SUM(InvoLine.Quantity)<> 0,
    SUM(InvoLine.Price)/ SUM(InvoLine.Quantity),
    SUM(InvoLine.Price)
  ) AS Price,
  IIf(
    SUM(InvoLine.Quantity)<> 0,
    SUM(Item.Cost * InvoLine.Quantity)/ SUM(InvoLine.Quantity),
    SUM(Item.Cost)
  ) AS ItemCost,
  IIf(
    SUM(InvoLine.Quantity) Mod SUM(Item.UnitsPerCase)<> 0,
    True,
    False
  ) AS Repack,
  0 AS LineNumber INTO PurchaseOrderItemDaysSupply
FROM
  (
    (
      (
        InvoLine
        INNER JOIN TranTypes ON InvoLine.TransType = TranTypes.Code
      )
      INNER JOIN Item ON InvoLine.Code = Item.ID
    )
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  )
  INNER JOIN SalesGroup ON SalesDepartment.SalesGroupID = SalesGroup.ID
WHERE
  (
    InvoLine.TransType = 'S'
    Or InvoLine.TransType = 'C'
  )
  And (
    InvoLine.SetupDate Between [parBeginDate]
    And [parEndDate]
  )
  And (
    Item.SalesDepartmentID Between [parBeginSalesDepartmentID]
    And [parEndSalesDepartmentID]
  )
  And (
    SalesGroup.Description = 'Liquor'
  )
GROUP BY
  Item.VendorID,
  Item.ID;
