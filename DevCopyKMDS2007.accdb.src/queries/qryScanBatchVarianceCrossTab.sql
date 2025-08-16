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
  Sum(
    IIf(
      Left(
        [ScanBatchLineCounts].[Location],
        1
      )= '1',
      [CountQuantity],
      0
    )
  ) AS Loc1,
  Sum(
    IIf(
      Left(
        [ScanBatchLineCounts].[Location],
        1
      )= '2',
      [CountQuantity],
      0
    )
  ) AS Loc2,
  Sum(
    IIf(
      Left(
        [ScanBatchLineCounts.Location],
        1
      )= '3',
      [CountQuantity],
      0
    )
  ) AS Loc3,
  Sum(
    IIf(
      Left(
        [ScanBatchLineCounts.Location],
        1
      )= '4',
      [CountQuantity],
      0
    )
  ) AS Loc4,
  vwItemsWithCurrentPricing.Price,
  Count(1) AS Scans,
  vwItemsWithCurrentPricing.Cost
FROM
  ScanBatchLineCounts
  INNER JOIN vwItemsWithCurrentPricing ON ScanBatchLineCounts.ItemID = vwItemsWithCurrentPricing.ID
GROUP BY
  ScanBatchLineCounts.ItemID,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.Price,
  vwItemsWithCurrentPricing.Cost;
