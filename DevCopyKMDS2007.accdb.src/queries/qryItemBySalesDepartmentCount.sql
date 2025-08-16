SELECT
  Item.SalesDepartmentID,
  Sum(1) AS [count]
FROM
  Item
WHERE
  (
    (
      (Item.SalesDepartmentID)> 2
      Or (Item.SalesDepartmentID)> 2
    )
  )
GROUP BY
  Item.SalesDepartmentID
HAVING
  (
    (
      (Item.SalesDepartmentID)> 1
    )
  );
