SELECT
  Format([ID], "00000000000") AS ID,
  Item.Onhand
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
  Format([ID], "00000000000");
