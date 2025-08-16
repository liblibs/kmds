SELECT
  OrderLine.OrdheadID AS PurchaseOrderID,
  OrderLine.LineID,
  OrderLine.ItemID,
  OrderLine.QuantityShipped,
  OrderLine.Price,
  0 AS QuantityScanned,
  0 AS FullCase INTO CompareScanItemToPurchaseOrder
FROM
  OrderLine
WHERE
  (
    (
      (OrderLine.OrdheadID)= [Forms]![ClubOrderCompare]![txtPurchaseOrderID_1]
    )
  );
