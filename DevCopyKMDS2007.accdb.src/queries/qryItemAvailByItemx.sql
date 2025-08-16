PARAMETERS parItemID Long;
SELECT
  Item.ID,
  [Onhand] + [TransferredToday] - [SoldTodayCr1] - [SoldTodayCr2] - [SoldTodayCr3] - [SoldTodayCr4] - [Reserve] AS Avail,
  Item.Onhand,
  IIf(
    IsNull([OnReserve]),
    0,
    [OnReserve]
  ) AS Reserve,
  Item.TransferredToday,
  Item.SoldTodayCr1,
  Item.SoldTodayCr2,
  Item.SoldTodayCr3,
  Item.SoldTodayCr4
FROM
  Item
  LEFT JOIN qryItemOnReserve ON Item.ID = qryItemOnReserve.ItemID
WHERE
  (
    (
      (Item.ID)= [parItemID]
    )
  );
