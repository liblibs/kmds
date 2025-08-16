SELECT
  ScanBatchLineCounts.ID,
  ScanBatchLineCounts.ItemID,
  vwItemsWithCurrentPricing.Description,
  ScanBatchLineCounts.SetupDate,
  ScanBatchLineCounts.CountQuantity,
  ScanBatchLineCounts.OnShelf AS OnShelfAtTimeOfScan,
  ScanBatchLineCounts.OnShelfDate,
  vwItemsWithCurrentPricing.Price,
  ScanBatchLineCounts.Location,
  vwItemsWithCurrentPricing.OnShelf,
  Cashier.Name
FROM
  (
    ScanBatchLineCounts
    INNER JOIN vwItemsWithCurrentPricing ON ScanBatchLineCounts.ItemID = vwItemsWithCurrentPricing.ID
  )
  INNER JOIN Cashier ON ScanBatchLineCounts.CashierID = Cashier.ID
ORDER BY
  ScanBatchLineCounts.ID;
