UPDATE
  PurchaseOrderItemWADS
SET
  PurchaseOrderItemWADS.SOQ = [SOQ] + [OrderMultiple]
WHERE
  (
    (
      (PurchaseOrderItemWADS.SOQ)> 0
      And (PurchaseOrderItemWADS.SOQ)<([SOQRaw] - 0.01)
    )
  );
