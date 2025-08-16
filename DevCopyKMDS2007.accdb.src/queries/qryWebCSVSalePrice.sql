SELECT
  ItemSalePrice.ItemID,
  ItemSalePrice.StartDate,
  ItemSalePrice.EndDate,
  ItemSalePrice.OldPrice,
  ItemSalePrice.SalePrice,
  ItemSalePrice.SetupDate,
  ItemSalePrice.ChangeDate
FROM
  ItemSalePrice
WHERE
  (
    (
      (ItemSalePrice.EndDate)> Now()
    )
  );
