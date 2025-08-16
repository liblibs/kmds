UPDATE
  Item
SET
  Item.OrderAt = [Onhand] - 1,
  Item.OrderTo = [Onhand],
  Item.UnitsPerOrder = [Onhand]
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
    )
    AND (
      (Item.Onhand)> 0
      And (Item.Onhand)< [UnitsPerCase]
    )
  );
