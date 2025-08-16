PARAMETERS parDaysSupplyGroupLevel Long,
parLowerAmount Long,
parUpperAmount Long;
UPDATE
  Item
  INNER JOIN DaysSupplyItemSales ON Item.ID = DaysSupplyItemSales.Code
SET
  Item.DaysSupplyGroupID = [parDaysSupplyGroupLevel]
WHERE
  (
    (
      (DaysSupplyItemSales.Sold) Between [parLowerAmount]
      And [parUpperAmount]
    )
  );
