INSERT INTO Audit (
  ItemID, PhysicalCount, OnShelf, Price,
  Cost, CountFront, SetupDate
)
SELECT
  Item.ID,
  qryMetrologicScanBatchVariance.Count AS PhysicalCount,
  qryMetrologicScanBatchVariance.OnShelf,
  Item.Price,
  Item.Cost,
  qryMetrologicScanBatchVariance.Count AS CountFront,
  Now() AS Expr1
FROM
  Item
  INNER JOIN qryMetrologicScanBatchVariance ON Item.ID = qryMetrologicScanBatchVariance.ID;
