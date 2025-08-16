SELECT
  PurchaseOrderLine.ItemID,
  Sum(
    PurchaseOrderLine.QuantityShipped
  ) AS OnOrder
FROM
  PurchaseOrderLine
GROUP BY
  PurchaseOrderLine.ItemID;
