SELECT
  UPCItem.ID,
  UPCItem.ItemID,
  Item.Description
FROM
  UPCItem
  INNER JOIN Item ON UPCItem.ItemID = Item.ID;
