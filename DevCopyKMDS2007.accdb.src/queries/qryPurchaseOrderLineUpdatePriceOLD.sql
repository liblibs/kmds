UPDATE
  PurchaseOrderLine
  INNER JOIN qryPurchaseOrderItemsWithTodaysPrice ON PurchaseOrderLine.ItemID = qryPurchaseOrderItemsWithTodaysPrice.ID
SET
  PurchaseOrderLine.Price = qryPurchaseOrderItemsWithTodaysPrice.sprice,
  PurchaseOrderLine.Cost = qryPurchaseOrderItemsWithTodaysPrice.cost
WHERE
  (
    (
      (PurchaseOrderLine.Price)<> qryPurchaseOrderItemsWithTodaysPrice.sprice
      Or (PurchaseOrderLine.Price) Is Null
    )
  )
  Or (
    (
      (PurchaseOrderLine.Cost)<> qryPurchaseOrderItemsWithTodaysPrice.Cost
      Or (PurchaseOrderLine.Cost) Is Null
    )
  );
