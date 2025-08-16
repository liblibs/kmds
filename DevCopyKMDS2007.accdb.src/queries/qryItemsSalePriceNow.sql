SELECT
  ItemSalePrice.ItemID,
  ItemSalePrice.StartDate,
  ItemSalePrice.EndDate,
  ItemSalePrice.SalePrice
FROM
  ItemSalePrice
WHERE
  (
    (
      (ItemSalePrice.StartDate)< Now()
    )
    AND (
      (ItemSalePrice.EndDate)> Now()
    )
  );
