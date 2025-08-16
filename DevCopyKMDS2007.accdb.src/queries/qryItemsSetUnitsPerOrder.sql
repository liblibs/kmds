UPDATE
  Item
SET
  Item.UnitsPerOrder = 1
WHERE
  (
    (
      (Item.UnitsPerOrder) Is Null
      Or (Item.UnitsPerOrder)< 1
    )
  );
