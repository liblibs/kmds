UPDATE
  Item
  INNER JOIN ItemsSoldToday ON Item.ID = ItemsSoldToday.Code
SET
  Item.Onhand = [Onhand] - [SoldToday],
  Item.DateLastSold = Now();
