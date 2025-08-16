SELECT
  PriceGroup.Description,
  Item.Description,
  Item.Price,
  Item.Cost,
  Item.Deposit,
  Item.UnitsPerCase,
  Item.QtyPack,
  Item.QtyPackPrice,
  PriceGroup.ID,
  Item.ID,
  IIf(
    [PRICE] > 0
    And [COST] > 0,
    ([PRICE] - [COST])/ [PRICE],
    0
  ) AS GP
FROM
  Item
  RIGHT JOIN PriceGroup ON Item.PriceGroupID = PriceGroup.ID
WHERE
  (
    (
      (PriceGroup.ID)<> 0
    )
  );
