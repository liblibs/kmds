INSERT INTO Audit (
  ItemID, OnShelf, PhysicalCount, VarianceAmount,
  Price, Cost
)
SELECT
  qryPSCAuditInventory.ID AS ItemID,
  qryPSCAuditInventory.OnShelf,
  IIf(
    [QuantityShipped] < 0, [OnShelf], [QuantityShipped]
  ) AS qs,
  (
    IIf(
      [QuantityShipped] < 0, [OnShelf], [QuantityShipped]
    )
  )- [onshelf] AS VarianceAmount,
  qryPSCAuditInventory.CurPrice,
  qryPSCAuditInventory.Cost
FROM
  qryPSCAuditInventory;
