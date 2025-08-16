SELECT
  PurchaseOrderLine.ItemID,
  Sum(
    PurchaseOrderLine.QuantityShipped
  ) AS OnOrder
FROM
  PurchaseOrderLine
WHERE
  (
    (
      (
        PurchaseOrderLine.PurchaseOrderID
      )<> [Forms]![PurchaseOrdersOnFile]![ID]
    )
  )
GROUP BY
  PurchaseOrderLine.ItemID;
