INSERT INTO PurchaseOrderLine (
  PurchaseOrderID, VendorID, LineID,
  ItemID, QuantityOrdered, QuantityShipped,
  Price, Cost
)
SELECT
  932 AS POID,
  1 AS VD,
  A1265.ID,
  A1265.Code,
  A1265.qty,
  A1265.qty,
  Item.Price,
  Item.Cost
FROM
  A1265
  INNER JOIN Item ON A1265.Code = Item.ID;
