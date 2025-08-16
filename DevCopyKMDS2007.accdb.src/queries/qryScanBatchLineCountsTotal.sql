SELECT
  ScanBatchLineCounts.ScanBatchHeaderID,
  ScanBatchLineCounts.ItemID,
  Sum(
    ScanBatchLineCounts.CountQuantity
  ) AS SumOfCountQuantity,
  Last(ScanBatchLineCounts.OnShelf) AS LastOfOnShelf
FROM
  ScanBatchLineCounts
GROUP BY
  ScanBatchLineCounts.ScanBatchHeaderID,
  ScanBatchLineCounts.ItemID;
