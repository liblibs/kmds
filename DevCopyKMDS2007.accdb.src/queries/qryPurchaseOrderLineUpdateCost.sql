UPDATE
  PurchaseOrderLine
  INNER JOIN PurchaseOrderLineUpdatePrice ON PurchaseOrderLine.ItemID = PurchaseOrderLineUpdatePrice.ID
SET
  PurchaseOrderLine.Cost = [PurchaseOrderLineUpdatePrice].[sprice]
WHERE
  (
    (
      (PurchaseOrderLine.Cost)<>(
        [PurchaseOrderLineUpdatePrice].[sprice]
      )
    )
    AND (
      (
        PurchaseOrderLineUpdatePrice.SalesDepartmentID
      )= 1
    )
  )
  OR (
    (
      (PurchaseOrderLine.Cost) Is Null
    )
    AND (
      (
        PurchaseOrderLineUpdatePrice.SalesDepartmentID
      )= 1
    )
  );
