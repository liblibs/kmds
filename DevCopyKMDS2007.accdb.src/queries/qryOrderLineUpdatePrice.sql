UPDATE
  OrderLine
  INNER JOIN qryPurchaseOrderItemsWithTodaysPrice ON OrderLine.ItemID = qryPurchaseOrderItemsWithTodaysPrice.ID
SET
  OrderLine.Price = qryPurchaseOrderItemsWithTodaysPrice.sprice,
  OrderLine.SalePrice = qryPurchaseOrderItemsWithTodaysPrice.Star
WHERE
  (
    (
      (OrderLine.Price)<> qryPurchaseOrderItemsWithTodaysPrice.sprice
      Or (OrderLine.Price) Is Null
    )
    And (
      (
        qryPurchaseOrderItemsWithTodaysPrice.SalesDepartmentID
      )= 1
    )
  );
