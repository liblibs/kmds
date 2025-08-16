UPDATE
  PurchaseOrderLine
  INNER JOIN qryPurchaseOrderItemsWithTodaysPrice ON PurchaseOrderLine.ItemID = qryPurchaseOrderItemsWithTodaysPrice.ID
SET
  PurchaseOrderLine.Price = qryPurchaseOrderItemsWithTodaysPrice.sprice
WHERE
  (
    (
      (PurchaseOrderLine.Price)<> [qryPurchaseOrderItemsWithTodaysPrice].[sprice]
      Or (PurchaseOrderLine.Price) Is Null
    )
  );
