UPDATE
  ScioInitialOrder
  INNER JOIN Item ON ScioInitialOrder.Code = Item.ID
SET
  Item.Price = [ScioInitialOrder].[price],
  Item.Cost = [ScioInitialOrder].[price]
WHERE
  (
    (
      (ScioInitialOrder.Price)<> [Item].[price]
    )
  );
