SELECT
  ScanBatchLineCounts.ItemID,
  Sum(
    ScanBatchLineCounts.CountQuantity
  ) AS SumOfCountQuantity,
  First(vwItemsWithCurrentPricing.Cost) AS FirstOfCost,
  First(
    vwItemsWithCurrentPricing.Price
  ) AS FirstOfPrice,
  vwItemsWithCurrentPricing.UnitsPerCase
FROM
  ScanBatchLineCounts
  INNER JOIN vwItemsWithCurrentPricing ON ScanBatchLineCounts.ItemID = vwItemsWithCurrentPricing.ID
GROUP BY
  ScanBatchLineCounts.ItemID,
  vwItemsWithCurrentPricing.UnitsPerCase
ORDER BY
  ScanBatchLineCounts.ItemID;
