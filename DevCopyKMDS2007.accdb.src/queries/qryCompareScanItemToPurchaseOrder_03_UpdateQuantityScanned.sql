UPDATE
  CompareScanItemToPurchaseOrderScanCounts
  INNER JOIN CompareScanItemToPurchaseOrder ON CompareScanItemToPurchaseOrderScanCounts.ItemID = CompareScanItemToPurchaseOrder.ItemID
SET
  CompareScanItemToPurchaseOrder.QuantityScanned = [SumOfCountQuantity];
