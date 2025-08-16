PARAMETERS parPurchaseOrderID Long;
SELECT
  PurchaseOrderLine.VendorID,
  PurchaseOrderLine.PurchaseOrderID,
  PurchaseOrderLine.LineID,
  PurchaseOrderLine.ItemID,
  PurchaseOrderLine.QuantityShipped,
  PurchaseOrderLine.Price,
  PurchaseOrderLine.Cost,
  PurchaseOrderLine.SetupDate,
  PurchaseOrderLine.ChangeDate,
  [QuantityOrdered] * [Price] AS ExtPrice,
  [QuantityOrdered] * [Cost] AS ExtCost
FROM
  PurchaseOrderLine
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
  )
ORDER BY
  PurchaseOrderLine.ItemID;
