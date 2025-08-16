UPDATE
  Item
SET
  Item.Cost = [Price]
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
    )
  );
