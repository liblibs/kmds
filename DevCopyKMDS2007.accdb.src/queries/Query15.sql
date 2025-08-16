UPDATE
  MeadowlandInventoryReport13
  INNER JOIN Item ON MeadowlandInventoryReport13.NewItemID = Item.ID
SET
  Item.OrderAt = [NewMin],
  Item.OrderTo = [NewMax];
