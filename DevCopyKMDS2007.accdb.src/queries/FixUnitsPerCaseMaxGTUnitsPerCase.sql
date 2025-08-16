UPDATE
  Item
SET
  Item.UnitsPerOrder = [UnitsPerCase]
WHERE
  (
    (
      (Item.OrderTo)> [UnitsPerCase]
    )
  );
