PARAMETERS parPurchaseOrderID Long;
UPDATE
  PurchaseOrderLine
  INNER JOIN Item ON PurchaseOrderLine.ItemID = Item.ID
SET
  Item.Onhand = [Onhand] + [QuantityShipped],
  Item.Cost = PurchaseOrderLine.cost
WHERE
  (
    (
      (
        PurchaseOrderLine.PurchaseOrderID
      )= [parPurchaseOrderID]
    )
    AND (
      (
        PurchaseOrderLine.QuantityShipped
      )<> 0
    )
  );
