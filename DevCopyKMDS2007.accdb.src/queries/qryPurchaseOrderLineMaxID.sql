PARAMETERS parPurchaseOrderID Long;
SELECT
  Max(PurchaseOrderLine.LineID) AS MaxLineID
FROM
  PurchaseOrderLine
WHERE
  (
    (
      (
        PurchaseOrderLine.PurchaseOrderID
      )= [parPurchaseOrderID]
    )
  );
