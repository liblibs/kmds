UPDATE
  Item
SET
  Item.QtyPack = 0
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
    )
  );
