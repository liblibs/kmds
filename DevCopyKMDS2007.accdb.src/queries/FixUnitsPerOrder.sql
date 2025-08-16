UPDATE
  Item
SET
  Item.UnitsPerOrder = [orderto] - [OrderAt]
WHERE
  (
    (
      (Item.OrderTo)> 0
      And (Item.OrderTo)< [UnitsPerCase]
    )
  );
