SELECT
  Vendor.ID AS VendorID,
  AverageDaySupply.ID AS AverageDaySupplyID,
  AverageDaySupplyLevels.ID AS AverageDaySupplyLevelsID,
  AverageDaySupplyLevels.DaysSupply,
  AverageDaySupplyLevels.MinUnitSales,
  AverageDaySupplyLevels.MaxUnitSales
FROM
  (
    Vendor
    INNER JOIN AverageDaySupply ON Vendor.AverageDaySupplyID = AverageDaySupply.ID
  )
  INNER JOIN AverageDaySupplyLevels ON AverageDaySupply.ID = AverageDaySupplyLevels.AverageDaySupplyID;
