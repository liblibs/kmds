SELECT
  PurchaseOrder.ID,
  PurchaseOrder.OrderDate,
  Vendor.ID,
  Vendor.Name,
  PurchaseOrderLine.ItemID,
  PurchaseOrderLine.QuantityShipped,
  PurchaseOrderLine.Price,
  PurchaseOrderLine.Cost,
  PurchaseOrderLine.SetupDate,
  PurchaseOrderLine.ChangeDate,
  [QuantityShipped] * [Price] AS ExtPrice,
  [QuantityShipped] * [Cost] AS ExtCost
FROM
  (
    PurchaseOrderLine
    INNER JOIN PurchaseOrder ON PurchaseOrderLine.PurchaseOrderID = PurchaseOrder.ID
  )
  INNER JOIN Vendor ON PurchaseOrder.VendorID = Vendor.ID
WHERE
  (
    (
      (
        PurchaseOrderLine.QuantityShipped
      )<> 0
    )
  );
