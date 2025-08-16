PARAMETERS parPurchaseOrderID Long;
UPDATE
  ItemLabels
  INNER JOIN PurchaseOrderLine ON ItemLabels.ID = PurchaseOrderLine.ItemID
SET
  ItemLabels.Labels = IIf(
    [PurchaseOrderLine].[QuantityShipped] <> 0,
    [PurchaseOrderLine].[QuantityShipped],
    1
  )
WHERE
  (
    (
      (
        PurchaseOrderLine.PurchaseOrderID
      )= [parPurchaseOrderID]
    )
    AND (
      (PurchaseOrderLine.PrintLabels)<> 0
    )
  );
