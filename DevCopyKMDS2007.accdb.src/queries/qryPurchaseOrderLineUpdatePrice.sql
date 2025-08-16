UPDATE
  PurchaseOrderLine
  INNER JOIN PurchaseOrderLineUpdatePrice ON PurchaseOrderLine.ItemID = PurchaseOrderLineUpdatePrice.ID
SET
  PurchaseOrderLine.Price = [PurchaseOrderLineUpdatePrice].[sprice]
WHERE
  (
    (
      (PurchaseOrderLine.Price)<> [PurchaseOrderLineUpdatePrice].[sprice]
      Or (PurchaseOrderLine.Price) Is Null
    )
  );
