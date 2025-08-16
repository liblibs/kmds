SELECT
  Item.ID,
  Item.AvgTurns
FROM
  Item
WHERE
  (
    (
      (Item.ID) Between 99900000000
      And 99999999999
    )
    AND (
      (Item.AvgTurns)> 0
    )
    AND (
      (Item.SalesDepartmentID)= 1
    )
  )
ORDER BY
  Item.ID;
