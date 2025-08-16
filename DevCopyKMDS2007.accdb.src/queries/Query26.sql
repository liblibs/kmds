INSERT INTO PurchaseOrderLine (
  PurchaseOrderID, VendorID, ItemID,
  QuantityOrdered, QuantityShipped,
  LineID
)
SELECT
  4113 AS PurchaseOrderID,
  1 AS VendorID,
  Sheet1.ItemID,
  Sheet1.qty AS QuantityOrdered,
  Sheet1.qty AS QuantityShipped,
  Sheet1.ID
FROM
  Sheet1
WHERE
  (
    (
      (Sheet1.qty)<> 0
    )
  );
