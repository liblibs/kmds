UPDATE
  Item
SET
  Item.DaysSupplyGroupID = 0,
  Item.AverageDailySales = 0
WHERE
  (
    (
      (Item.DaysSupplyGroupID)<> 0
      Or (Item.DaysSupplyGroupID) Is Null
    )
  )
  OR (
    (
      (Item.AverageDailySales)<> 0
      Or (Item.AverageDailySales) Is Null
    )
  );
