SELECT
  CompareScanItemToPurchaseOrder.PurchaseOrderID,
  CompareScanItemToPurchaseOrder.LineID,
  CompareScanItemToPurchaseOrder.ItemID,
  Item.Description,
  CompareScanItemToPurchaseOrder.QuantityShipped,
  CompareScanItemToPurchaseOrder.QuantityScanned,
  [QuantityScanned] - [QuantityShipped] AS Diff,
  Item.Price,
  CompareScanItemToPurchaseOrder.FullCase,
  [QuantityShipped] * [Item].[Price] AS ExtShipped,
  [QuantityScanned] * [Item].[Price] AS ExtScanned,
  (
    [QuantityScanned] - [QuantityShipped]
  )* [Item].[Price] AS ExtDiff
FROM
  CompareScanItemToPurchaseOrder
  INNER JOIN Item ON CompareScanItemToPurchaseOrder.ItemID = Item.ID;
