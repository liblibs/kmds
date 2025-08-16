UPDATE
  Item
SET
  Item.RestockAmount = 0
WHERE
  (
    (
      (Item.RestockAmount)<> 0
    )
  );
