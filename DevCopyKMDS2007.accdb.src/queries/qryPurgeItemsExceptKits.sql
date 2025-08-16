UPDATE
  Item
  INNER JOIN KitItem ON Item.ID = KitItem.ItemID
SET
  Item.AvgTurns = 0
WHERE
  (
    (
      (Item.AvgTurns)= 999
    )
  );
