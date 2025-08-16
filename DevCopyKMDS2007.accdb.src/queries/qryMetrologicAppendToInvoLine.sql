INSERT INTO InvoLine (
  InvoHeadID, ID, CashierID, Code, SalesDept,
  Quantity, Price, Cost, SetupDate,
  InputCode, TransType, CashSysID
)
SELECT
  Format(
    Now(),
    "yymmddhhnn"
  ) AS Expr1,
  qryMetrologicScanBatchVariance.InvoLineID,
  qryMetrologicScanBatchVariance.CashierID,
  Item.ID,
  qryMetrologicScanBatchVariance.OnShelf,
  qryMetrologicScanBatchVariance.Count AS PhysicalCount,
  Item.Price,
  Item.Cost,
  Now() AS Expr3,
  "Metrologic" AS InputCode,
  "Y" AS TransType,
  50 AS Expr2
FROM
  qryMetrologicScanBatchVariance
  INNER JOIN Item ON qryMetrologicScanBatchVariance.ID = Item.ID;
