SELECT
  Item.VendorID,
  Sum(1) AS [count]
FROM
  Item
WHERE
  (
    (
      (Item.SalesDepartmentID)> 2
      Or (Item.SalesDepartmentID)> 2
    )
    AND (
      (Item.Price)> 0
    )
  )
GROUP BY
  Item.VendorID
HAVING
  (
    (
      (Item.VendorID)> 1
    )
  );
