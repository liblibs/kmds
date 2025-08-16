SELECT
  Item.SalesDepartmentID,
  Item.ID,
  Item.Onhand,
  ResendMonthEnd.delta,
  [Onhand] + IIf(
    IsNull([delta]),
    0,
    [delta]
  ) AS MonthEndOnhand
FROM
  Item
  LEFT JOIN ResendMonthEnd ON Item.ID = ResendMonthEnd.Code
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
    )
    AND (
      (
        [Onhand] + IIf(
          IsNull([delta]),
          0,
          [delta]
        )
      )> 0
    )
  )
ORDER BY
  Item.ID;
