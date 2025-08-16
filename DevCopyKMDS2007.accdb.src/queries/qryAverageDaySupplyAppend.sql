INSERT INTO AverageDaySupplyLevels (
  GroupLevel, DaysSupply, MinUnitSales,
  MaxUnitSales, AverageDaySupplyID
)
SELECT
  AverageDaySupplyLevelsTemplate.GroupLevel,
  AverageDaySupplyLevelsTemplate.DaysSupply,
  AverageDaySupplyLevelsTemplate.MinUnitSales,
  AverageDaySupplyLevelsTemplate.MaxUnitSales,
  Forms!AverageDaySupply!ID AS ADSID
FROM
  AverageDaySupplyLevelsTemplate;
