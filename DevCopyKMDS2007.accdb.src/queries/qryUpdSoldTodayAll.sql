PARAMETERS parRegister Short,
parSold Short,
parItemID Long;
UPDATE
  Item
SET
  Item.SoldTodayCr1 = IIf(
    [parRegister] = 1, [parSold], [SoldTodayCr1]
  ),
  Item.SoldTodayCr2 = IIf(
    [parRegister] = 2, [parSold], [SoldTodayCr2]
  ),
  Item.SoldTodayCr3 = IIf(
    [parRegister] = 3, [parSold], [SoldTodayCr3]
  ),
  Item.SoldTodayCr4 = IIf(
    [parRegister] = 4, [parSold], [SoldTodayCr4]
  )
WHERE
  (
    (
      (Item.ID)= [parItemID]
    )
  );
