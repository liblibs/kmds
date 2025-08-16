SELECT
  ScanBatchLineCounts.ItemID,
  Sum(
    ScanBatchLineCounts.CountQuantity
  ) AS CountQ,
  Last(ScanBatchLineCounts.OnShelf) AS OS,
  Count(
    ScanBatchLineCounts.CountQuantity
  ) AS ScanQ
FROM
  ScanBatchLineCounts
GROUP BY
  ScanBatchLineCounts.ItemID
HAVING
  (
    (
      (ScanBatchLineCounts.ItemID)= 1402
    )
  );
