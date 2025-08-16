SELECT
  Item.SalesDepartmentID,
  PurchaseOrderLine.Price,
  Item.Price
FROM
  PurchaseOrderLine
  INNER JOIN Item ON PurchaseOrderLine.ItemID = Item.ID
WHERE
  (
    (
      (Item.SalesDepartmentID)<> 1
    )
    And (
      (Item.Price)<> PurchaseOrderLine.price
    )
  );
