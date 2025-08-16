PARAMETERS parDaysSupplyGroupLevel Long,
parLowerAmount Long,
parUpperAmount Long,
parVendorID Long;
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
    AND (
      (Item.VendorID)= [parVendorID]
    )
  );
