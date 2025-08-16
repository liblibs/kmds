INSERT INTO InvoLine (
  CashierID, Code, SalesDept, Quantity,
  Price, Cost, SetupDate, InputCode,
  TransType, InvoHeadID, CashSysID
)
SELECT
  ScanBatchLineCounts.CashierID,
  ScanBatchLineCounts.ItemID,
  ScanBatchLineCounts.OnShelf,
  ScanBatchLineCounts.CountQuantity AS PhysicalCount,
  Item.Price,
  Item.Cost,
  ScanBatchLineCounts.ChangeDate,
  "ScanBatch" AS InputCode,
  "Y" AS TransType,
  Format(
    Now(),
    "yymmddhhnn"
  ) AS Expr1,
  50 AS Expr2
FROM
  (
    ScanBatchLineCounts
    INNER JOIN Item ON ScanBatchLineCounts.ItemID = Item.ID
  )
  INNER JOIN qryMetrologicScanBatchVariance ON Item.ID = qryMetrologicScanBatchVariance.ID;
