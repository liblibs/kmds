SELECT
  Item.ID,
  Item.Onhand,
  Format([ID], "00000000000") AS Code11
FROM
  Item
WHERE
  (
    (
      (Item.Onhand)> 0
    )
    AND (
      (Item.SalesDepartmentID)= 1
    )
  )
ORDER BY
  Item.ID;
