INSERT INTO Audit (
  ItemID, PhysicalCount, OnShelf, Price,
  Cost, CountFront, SetupDate, ChangeDate
)
SELECT
  ScanBatchLineCounts.ItemID,
  Sum(
    ScanBatchLineCounts.CountQuantity
  ) AS PhysicalCount,
  Last(ScanBatchLineCounts.OnShelf) AS LastOfOnShelf,
  First(Item.Price) AS FirstOfPrice,
  First(Item.Cost) AS FirstOfCost,
  Sum(
    ScanBatchLineCounts.CountQuantity
  ) AS CountFront,
  First(ScanBatchLineCounts.SetupDate) AS FirstOfSetupDate,
  Last(ScanBatchLineCounts.ChangeDate) AS LastOfChangeDate
FROM
  ScanBatchLineCounts
  INNER JOIN Item ON ScanBatchLineCounts.ItemID = Item.ID
GROUP BY
  ScanBatchLineCounts.ItemID;
