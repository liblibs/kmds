UPDATE
  CompareScanItemToPurchaseOrder
  INNER JOIN Item ON CompareScanItemToPurchaseOrder.ItemID = Item.ID
SET
  CompareScanItemToPurchaseOrder.FullCase = True
WHERE
  (
    (
      (
        CompareScanItemToPurchaseOrder.QuantityShipped
      )>= [UnitsPerCase]
    )
  );
