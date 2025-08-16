SELECT
  Item.SalesDepartmentID,
  Item.ID,
  Format([id], "000000") AS NewUPC
FROM
  Item
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
    )
  );
