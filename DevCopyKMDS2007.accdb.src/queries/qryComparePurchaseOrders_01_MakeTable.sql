SELECT
  PurchaseOrderLine.PurchaseOrderID,
  PurchaseOrderLine.LineID,
  PurchaseOrderLine.ItemID,
  PurchaseOrderLine.QuantityShipped,
  PurchaseOrderLine.Price,
  0 AS QuantityScanned,
  0 AS FullCase INTO CompareScanItemToPurchaseOrder
FROM
  PurchaseOrderLine
WHERE
  (
    (
      (
        PurchaseOrderLine.PurchaseOrderID
      )= [Forms]![PurchaseOrderCompare]![txtPurchaseOrderID_1]
    )
  );
