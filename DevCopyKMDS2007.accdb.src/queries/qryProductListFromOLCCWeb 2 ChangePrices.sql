UPDATE
  Item
  INNER JOIN qryProductListFromOLCCWebsite ON Item.ID = qryProductListFromOLCCWebsite.Code
SET
  Item.Price = qryProductListFromOLCCWebsite.price,
  Item.ChangeDate = Now()
WHERE
  (
    (
      (
        qryProductListFromOLCCWebsite.Price
      )<> item.price
    )
  );
