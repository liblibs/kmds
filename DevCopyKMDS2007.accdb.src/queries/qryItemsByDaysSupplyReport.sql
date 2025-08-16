SELECT
  SalesDepartment.ID AS SaleDpt,
  SalesDepartment.Description AS SaleDptDesc,
  Item.ID,
  Item.Description,
  Item.Price,
  Item.Onhand,
  qryItemsByDaysSupplyAvgDailySales.AverageDailySales,
  CLng(
    IIf(
      [AverageDailySales] > 0, [Onhand] / [AverageDailySales],
      0
    )
  ) AS DaysSupply,
  Forms!PrintInventoryReport!txtNumericRangeFrom AS RangeOfDaysFrom,
  Forms!PrintInventoryReport!txtNumericRangeTo AS RangeOfDaysTo,
  Item.OrderAt,
  Item.OrderTo,
  qryItemsByDaysSupplyAvgDailySales.Sold,
  qryItemsByDaysSupplyAvgDailySales.Days
FROM
  (
    SalesDepartment
    INNER JOIN Item ON SalesDepartment.ID = Item.SalesDepartmentID
  )
  INNER JOIN qryItemsByDaysSupplyAvgDailySales ON Item.ID = qryItemsByDaysSupplyAvgDailySales.ID
WHERE
  (
    (
      (Item.Onhand)<> 0
    )
    And (
      (
        CLng(
          IIf(
            [AverageDailySales] > 0, [Onhand] / [AverageDailySales],
            0
          )
        )
      ) Between CLng(
        Forms!PrintInventoryReport!txtNumericRangeFrom
      )
      And CLng(
        Forms!PrintInventoryReport!txtNumericRangeTo
      )
    )
    And (
      (
        IIf(
          Forms!PrintInventoryReport!chkPrintZeroOnhands,
          -1, 0
        )
      )= 0
    )
  )
  Or (
    (
      (
        CLng(
          IIf(
            [AverageDailySales] > 0, [Onhand] / [AverageDailySales],
            0
          )
        )
      ) Between CLng(
        Forms!PrintInventoryReport!txtNumericRangeFrom
      )
      And CLng(
        Forms!PrintInventoryReport!txtNumericRangeTo
      )
    )
    And (
      (
        IIf(
          Forms!PrintInventoryReport!chkPrintZeroOnhands,
          -1, 0
        )
      )=-1
    )
  )
ORDER BY
  SalesDepartment.ID,
  Item.ID;
