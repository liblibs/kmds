SELECT
  Item.SalesDepartmentID,
  Sum(1) AS [count]
FROM
  Item
WHERE
  (
    (
      (Item.Price)> 0
    )
  )
GROUP BY
  Item.SalesDepartmentID;
