SELECT
  PurchaseOrderLine.PurchaseOrderID,
  OLCCEPS.OrderNumber,
  OLCCEPS.OLCCItem,
  OLCCEPS.Description,
  OLCCEPS.Price,
  PurchaseOrderLine.ItemID,
  PurchaseOrderLine.QuantityOrdered
FROM
  OLCCEPS
  LEFT JOIN PurchaseOrderLine ON OLCCEPS.OLCCItem = PurchaseOrderLine.ItemID
WHERE
  (
    (
      (
        PurchaseOrderLine.PurchaseOrderID
      )= [PO]
    )
    AND (
      (OLCCEPS.OrderNumber)= [EPS]
    )
  );
