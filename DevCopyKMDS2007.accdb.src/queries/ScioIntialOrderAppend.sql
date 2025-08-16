INSERT INTO PurchaseOrderLine (
  PurchaseOrderID, LineID, VendorID,
  QuantityOrdered, QuantityShipped,
  Price, Cost, ItemID
)
SELECT
  [PurchaseOrderNumber] AS PurchaseOrderID,
  ScioInitialOrder.ID AS LineID,
  1 AS VendorID,
  ScioInitialOrder.OrderQty AS QuantityOrdered,
  ScioInitialOrder.OrderQty AS QuantityShipped,
  ScioInitialOrder.Price,
  ScioInitialOrder.Price AS Cost,
  ScioInitialOrder.Code
FROM
  ScioInitialOrder
WHERE
  (
    (
      (ScioInitialOrder.ID) Is Not Null
    )
  );
