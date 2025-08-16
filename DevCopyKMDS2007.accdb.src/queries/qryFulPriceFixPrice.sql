UPDATE
  Fulprice
  INNER JOIN Item ON Fulprice.Code = Item.ID
SET
  Item.Price = (fulprice.Price * 0.01)
WHERE
  (
    (
      (Item.Price)<>(fulprice.Price * 0.01)
    )
  );
