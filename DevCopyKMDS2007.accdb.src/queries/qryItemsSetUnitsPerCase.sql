UPDATE
  Item
SET
  Item.UnitsPerCase = 1
WHERE
  (
    (
      (Item.UnitsPerCase) Is Null
      Or (Item.UnitsPerCase)< 1
    )
  );
