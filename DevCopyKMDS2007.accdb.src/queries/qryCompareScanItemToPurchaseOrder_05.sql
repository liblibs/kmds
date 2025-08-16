SELECT
  CompareScanItemToPurchaseOrder.PurchaseOrderID,
  CompareScanItemToPurchaseOrder.LineID,
  CompareScanItemToPurchaseOrder.ItemID,
  Item.Description,
  CompareScanItemToPurchaseOrder.QuantityShipped,
  CompareScanItemToPurchaseOrder.QuantityScanned,
  [QuantityScanned] - [QuantityShipped] AS Diff,
  Item.Price,
  CompareScanItemToPurchaseOrder.FullCase
FROM
  CompareScanItemToPurchaseOrder
  INNER JOIN Item ON CompareScanItemToPurchaseOrder.ItemID = Item.ID;
