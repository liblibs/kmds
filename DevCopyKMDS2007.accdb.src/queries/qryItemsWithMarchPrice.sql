SELECT
  Item.ID,
  Item.SalesDepartmentID,
  Item.Price,
  ItemSalePrice.StartDate,
  ItemSalePrice.EndDate,
  ItemSalePrice.SalePrice,
  IIf(
    [StartDate] > #2/28/2005#, [SalePrice],
    [Price]
  ) AS SPrice
FROM
  Item
  LEFT JOIN ItemSalePrice ON Item.ID = ItemSalePrice.ItemID
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
    )
  );
