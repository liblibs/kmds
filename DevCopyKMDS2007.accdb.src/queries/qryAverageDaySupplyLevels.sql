SELECT
  AverageDaySupplyLevels.AverageDaySupplyID,
  AverageDaySupplyLevels.GroupLevel,
  AverageDaySupplyLevels.DaysSupply,
  AverageDaySupplyLevels.MinUnitSales,
  AverageDaySupplyLevels.MaxUnitSales,
  AverageDaySupplyLevels.SetupDate,
  AverageDaySupplyLevels.ChangeDate
FROM
  AverageDaySupplyLevels
ORDER BY
  AverageDaySupplyLevels.GroupLevel;
