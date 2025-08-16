INSERT INTO UPCItem (ID, ItemID)
SELECT
  Item.ID,
  Item.ID AS ItemID
FROM
  Item;
