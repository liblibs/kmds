INSERT INTO Audit (
  ItemID, VarianceAmount, Price, Cost,
  OnShelf, CountFront, PhysicalCount
)
SELECT
  OLCCAuditVariance.ItemID,
  OLCCAuditVariance.Variance,
  vwItemsWithCurrentPricing.Price,
  vwItemsWithCurrentPricing.Cost,
  vwItemsWithCurrentPricing.OnShelf,
  [OnShelf] + [Variance] AS Expr1,
  [OnShelf] + [Variance] AS Expr2
FROM
  OLCCAuditVariance
  INNER JOIN vwItemsWithCurrentPricing ON OLCCAuditVariance.ItemId = vwItemsWithCurrentPricing.ID;
