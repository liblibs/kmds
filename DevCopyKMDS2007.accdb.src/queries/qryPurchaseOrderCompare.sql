SELECT
  PurchaseOrderLine.PurchaseOrderID,
  PurchaseOrderLine.ItemID,
  Item.Description,
  PurchaseOrderLine.QuantityShipped,
  PurchaseOrderLine.Price
FROM
  PurchaseOrderLine
  INNER JOIN Item ON PurchaseOrderLine.ItemID = Item.ID
WHERE
  (
    (
      (
        PurchaseOrderLine.PurchaseOrderID
      )= [Forms]![PurchaseOrderCreate]![txtSelectedInvoice]
    )
  );
