UPDATE
  RecreateInventoryTransmit
  INNER JOIN Item ON RecreateInventoryTransmit.Code = Item.ID
SET
  Item.AvgTurns = [AvgTurns] + [RecreateInventoryTransmit].[movement];
