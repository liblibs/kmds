SELECT
  PurchaseOrderLine.PurchaseOrderID,
  Sum(
    PurchaseOrderLine.QuantityShipped
  ) AS Quantity,
  Sum([QuantityShipped] * [Price]) AS Dollars
FROM
  PurchaseOrderLine
GROUP BY
  PurchaseOrderLine.PurchaseOrderID;
