SELECT
  ScanBatchLineCounts.ID,
  ScanBatchLineCounts.ScanBatchHeaderID,
  ScanBatchLineCounts.ItemID,
  ScanBatchLineCounts.InputMethod,
  ScanBatchLineCounts.CashierID,
  ScanBatchLineCounts.Location,
  ScanBatchLineCounts.ScanCode,
  ScanBatchLineCounts.CountQuantity,
  ScanBatchLineCounts.OnShelf,
  ScanBatchLineCounts.SetupDate,
  ScanBatchLineCounts.ChangeDate,
  Item.Description,
  Item.Price
FROM
  ScanBatchLineCounts
  INNER JOIN Item ON ScanBatchLineCounts.ItemID = Item.ID;
