INSERT INTO CompareScanItemToPurchaseOrder (
  ItemID, QuantityShipped, Price, PurchaseOrderID,
  QuantityScanned, FullCase
)
SELECT
  CompareScanItemToPurchaseOrderScanCounts.ItemID,
  0 AS QuantityShipped,
  Item.Price,
  -1 AS PurchaseOrderID,
  CompareScanItemToPurchaseOrderScanCounts.SumOfCountQuantity,
  0 AS FullCase
FROM
  (
    CompareScanItemToPurchaseOrderScanCounts
    LEFT JOIN CompareScanItemToPurchaseOrder ON CompareScanItemToPurchaseOrderScanCounts.ItemID = CompareScanItemToPurchaseOrder.ItemID
  )
  INNER JOIN Item ON CompareScanItemToPurchaseOrderScanCounts.ItemID = Item.ID
WHERE
  (
    (
      (
        CompareScanItemToPurchaseOrder.ItemID
      ) Is Null
    )
  );
