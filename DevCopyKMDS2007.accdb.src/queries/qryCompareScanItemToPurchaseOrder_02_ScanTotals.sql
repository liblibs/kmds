SELECT
  ScanBatchLineCounts.ItemID,
  Sum(
    ScanBatchLineCounts.CountQuantity
  ) AS SumOfCountQuantity INTO CompareScanItemToPurchaseOrderScanCounts
FROM
  ScanBatchLineCounts
GROUP BY
  ScanBatchLineCounts.ItemID;
