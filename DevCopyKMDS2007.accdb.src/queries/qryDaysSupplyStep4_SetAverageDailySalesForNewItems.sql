UPDATE
  Item
  INNER JOIN DaysSupplyADSForNewItems ON Item.ID = DaysSupplyADSForNewItems.Code
SET
  Item.AverageDailySales = DaysSupplyADSForNewItems.ADS;
