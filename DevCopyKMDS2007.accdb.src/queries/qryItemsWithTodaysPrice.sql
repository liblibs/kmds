SELECT
  Item.ID,
  Item.Price,
  qryItemsSalePriceNow.SalePrice,
  IIf(
    IsNull([StartDate]),
    [Price],
    [SalePrice]
  ) AS DP,
  qryItemsSalePriceNow.StartDate,
  qryItemsSalePriceNow.EndDate
FROM
  Item
  LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID;
