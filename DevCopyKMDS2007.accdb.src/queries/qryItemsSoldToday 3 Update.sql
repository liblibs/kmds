UPDATE
  Item
  INNER JOIN ItemsSoldToday ON Item.ID = ItemsSoldToday.Code
SET
  Item.SoldTodayCr1 = IIf(
    ([CashSysID] = 1),
    [SumOfQuantity],
    [SoldTodayCR1]
  ),
  Item.SoldTodayCr2 = IIf(
    ([CashSysID] = 2),
    [SumOfQuantity],
    [SoldTodayCR2]
  ),
  Item.SoldTodayCr3 = IIf(
    ([CashSysID] = 3),
    [SumOfQuantity],
    [SoldTodayCR3]
  ),
  Item.SoldTodayCr4 = IIf(
    ([CashSysID] = 4),
    [SumOfQuantity],
    [SoldTodayCR4]
  );
