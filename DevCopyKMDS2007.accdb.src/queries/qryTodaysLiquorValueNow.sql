SELECT
  Sum(
    IIf(
      IsNull(qryItemsSalePriceNow.saleprice),
      [Price],
      qryItemsSalePriceNow.SalePrice
    )* [onhand]
  ) AS CurrentLiquorValue
FROM
  Item
  LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
    )
  );
