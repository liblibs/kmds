SELECT
  vwItemsWithCurrentPricing.ID,
  IIf(
    [AverageDailySales] = 0, 0, [OnShelf] / [AverageDailySales]
  ) AS CurDaysSupply,
  vwItemsWithCurrentPricing.OnShelf,
  vwItemsWithCurrentPricing.AverageDailySales
FROM
  vwItemsWithCurrentPricing;
