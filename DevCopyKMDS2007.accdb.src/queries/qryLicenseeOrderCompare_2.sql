SELECT
  OrderLine.OrdheadID,
  OrderLine.ItemID,
  Item.Description,
  OrderLine.QuantityShipped,
  OrderLine.Price
FROM
  OrderLine
  INNER JOIN Item ON OrderLine.ItemID = Item.ID
WHERE
  (
    (
      (OrderLine.OrdheadID)= [Forms]![ClubOrderCompare]![txtPurchaseOrderID_2]
    )
  );
