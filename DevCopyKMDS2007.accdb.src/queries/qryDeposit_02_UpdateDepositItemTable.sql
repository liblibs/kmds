UPDATE
  DepositTable
  INNER JOIN Item ON DepositTable.DepositItemID = Item.ID
SET
  Item.Onhand = [DepositTotal] / [Item].[price];
