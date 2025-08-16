UPDATE
  KenoSalesHist
  INNER JOIN Item ON KenoSalesHist.ItemID = Item.ID
SET
  Item.UnitsPerCase = [ucase],
  Item.UnitsPerOrder = [uorder];
