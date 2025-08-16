INSERT INTO PurchaseOrderLine (
  PurchaseOrderID, VendorID, LineID,
  ItemID, QuantityOrdered, QuantityShipped
)
SELECT
  2279 AS PurchaseOrderID,
  1 AS VendorID,
  Sheet1.id AS LineID,
  Sheet1.field1 AS ItemID,
  Sheet1.field2 AS QuantityOrdered,
  Sheet1.Field2 AS Expr1
FROM
  Sheet1;
