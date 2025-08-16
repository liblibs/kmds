PARAMETERS parRegister Short,
parSold Short,
parItemID Long,
parTransType Text (255);
UPDATE
  Item
SET
  Item.SoldTodayCr1 = [SoldTodayCr1] + IIf([parRegister] = 1, [parSold], 0),
  Item.SoldTodayCr2 = [SoldTodayCr2] + IIf([parRegister] = 2, [parSold], 0),
  Item.SoldTodayCr3 = [SoldTodayCr3] + IIf([parRegister] = 3, [parSold], 0),
  Item.SoldTodayCr4 = [SoldTodayCr4] + IIf([parRegister] = 4, [parSold], 0),
  Item.DateLastSold = Now(),
  Item.RestockAmount = IIf(
    [parTransType] <> "C", [RestockAmount] + [parSold],
    [RestockAmount]
  )
WHERE
  (
    (
      (Item.ID)= [parItemID]
    )
  );
