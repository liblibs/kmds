PARAMETERS parItemID Long,
parSold Short;
UPDATE
  Item
SET
  Item.TransferredToday = [TransferredToday] + [parSold]
WHERE
  (
    (
      (Item.ID)= [parItemID]
    )
  );
