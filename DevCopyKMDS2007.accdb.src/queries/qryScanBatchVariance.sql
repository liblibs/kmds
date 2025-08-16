SELECT
  ScanBatchLineCounts.ItemID,
  vwItemsWithCurrentPricing.Description,
  First(ScanBatchLineCounts.SetupDate) AS FirstScan,
  Last(ScanBatchLineCounts.SetupDate) AS LastScan,
  Sum(
    ScanBatchLineCounts.CountQuantity
  ) AS SumOfCountQuantity,
  Last(ScanBatchLineCounts.OnShelf) AS OnShelfAtTimeOfScan,
  First(
    vwItemsWithCurrentPricing.OnShelf
  ) AS OnShelfAtTimeOfScanNOW,
  Last(
    ScanBatchLineCounts.OnShelfDate
  ) AS LastOfOnShelfDate,
  ScanBatchLineCounts.Location,
  First(
    vwItemsWithCurrentPricing.Price
  ) AS FirstOfPrice,
  First(vwItemsWithCurrentPricing.Cost) AS FirstOfCost
FROM
  ScanBatchLineCounts
  INNER JOIN vwItemsWithCurrentPricing ON ScanBatchLineCounts.ItemID = vwItemsWithCurrentPricing.ID
GROUP BY
  ScanBatchLineCounts.ItemID,
  vwItemsWithCurrentPricing.Description,
  ScanBatchLineCounts.Location;
