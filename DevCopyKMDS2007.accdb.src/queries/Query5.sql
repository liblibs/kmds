INSERT INTO PurchaseOrderLine (
  ItemID, Price, VendorID, Cost, PurchaseOrderID,
  LineID, QuantityOrdered, QuantityShipped
)
SELECT
  TalentPO.ItemID,
  TalentPO.Price,
  TalentPO.VendorID,
  TalentPO.Cost,
  TalentPO.PurchaseOrderID,
  TalentPO.LineID,
  Sheet1.Qty,
  Sheet1.Qty
FROM
  TalentPO
  INNER JOIN Sheet1 ON TalentPO.ItemID = Sheet1.[Item Code];
