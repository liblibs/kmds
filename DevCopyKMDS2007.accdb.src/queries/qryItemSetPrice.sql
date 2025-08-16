PARAMETERS parItemID Long,
parPrice Currency;
UPDATE
  Item
SET
  Item.Price = [parPrice]
WHERE
  (
    (
      (Item.ID)= [parItemID]
    )
  );
