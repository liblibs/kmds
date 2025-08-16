SELECT
  Item.ID,
  IIf(
    IsNull([SumOfQuantity]),
    0,
    [SumOfQuantity]
  ) AS Sold,
  DateDiff(
    "d", Forms!PrintInventoryReport!txtDateRangeFrom,
    Forms!PrintInventoryReport!txtDateRangeTo
  )+ 1 AS Days,
  IIf([Days] = 0, 0, [Sold] / [Days]) AS AverageDailySales
FROM
  SalesDepartment
  INNER JOIN (
    Item
    LEFT JOIN qryItemsByDaysSupplySoldTotals ON Item.ID = qryItemsByDaysSupplySoldTotals.Code
  ) ON SalesDepartment.ID = Item.SalesDepartmentID
WHERE
  (
    (
      (SalesDepartment.ID) Between [Forms]![PrintInventoryReport]![cboSalesDepartmentIDBegin]
      And [Forms]![PrintInventoryReport]![cboSalesDepartmentIDEnd]
    )
  );
