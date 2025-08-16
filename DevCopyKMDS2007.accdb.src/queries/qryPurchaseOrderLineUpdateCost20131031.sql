UPDATE
  PurchaseOrderLine
  INNER JOIN qryPurchaseOrderItemsWithTodaysPrice ON PurchaseOrderLine.ItemID = qryPurchaseOrderItemsWithTodaysPrice.ID
SET
  PurchaseOrderLine.Cost = qryPurchaseOrderItemsWithTodaysPrice.sprice
WHERE
  (
    (
      (PurchaseOrderLine.Cost)<>(
        qryPurchaseOrderItemsWithTodaysPrice.sprice
      )
    )
    And (
      (
        qryPurchaseOrderItemsWithTodaysPrice.SalesDepartmentID
      )= 1
    )
  )
  Or (
    (
      (PurchaseOrderLine.Cost) Is Null
    )
    And (
      (
        qryPurchaseOrderItemsWithTodaysPrice.SalesDepartmentID
      )= 1
    )
  );
