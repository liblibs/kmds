UPDATE
  Item
SET
  Item.OrderAt = [Onhand] - 1,
  Item.OrderTo = [Onhand],
  Item.UnitsPerOrder = [UnitsPerCase]
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
    )
    AND (
      (Item.Onhand)>= [UnitsPerCase]
    )
  );
