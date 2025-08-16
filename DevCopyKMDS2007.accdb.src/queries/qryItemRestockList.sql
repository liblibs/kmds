SELECT
  IIf(
    [SalesdepartmentID] = 1,
    ([item].[id] - 99900000000),
    [item].[id]
  ) AS ID,
  Item.Description,
  [Onhand] - [SoldTodayCr1] - [SoldTodayCr2] - [SoldTodayCr3] - [SoldTodayCr4] + [TransferredToday] AS OH,
  IIf(
    IsNull([OnReserve]),
    0,
    [OnReserve]
  ) AS Reserved,
  Item.RestockAmount,
  SalesDepartment.IncludeInInventoryValueReports
FROM
  (
    Item
    LEFT JOIN qryItemOnReserve ON Item.ID = qryItemOnReserve.ItemID
  )
  INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
WHERE
  (
    (
      (Item.RestockAmount)<> 0
    )
    AND (
      (
        SalesDepartment.IncludeInInventoryValueReports
      )<> 0
    )
  )
ORDER BY
  Item.SalesDepartmentID,
  Item.ID;
