SELECT
  SalesGroup.GroupLevel,
  SalesGroup.ID,
  SalesGroup.Description,
  SalesDepartment.ID,
  SalesDepartment.Description,
  SalesDepartment.Discounts,
  SalesDepartment.NegativeOnhand,
  SalesDepartment.AllowDiscounts,
  SalesDepartment.IncludeInInventoryValueReports,
  SalesDepartment.TargetGP,
  Item.ID,
  Item.Description,
  IIf(
    IsNull(
      [qryItemsBySalesDepartmentTotals].[count]
    ),
    0,
    [qryItemsBySalesDepartmentTotals].[count]
  ) AS Items
FROM
  (
    (
      SalesGroup
      INNER JOIN SalesDepartment ON SalesGroup.ID = SalesDepartment.SalesGroupID
    )
    LEFT JOIN Item ON SalesDepartment.ID = Item.ID
  )
  LEFT JOIN qryItemsBySalesDepartmentTotals ON SalesDepartment.ID = qryItemsBySalesDepartmentTotals.SalesDepartmentID
ORDER BY
  SalesGroup.GroupLevel,
  SalesGroup.ID,
  SalesDepartment.ID;
