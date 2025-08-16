SELECT
  Item.ID,
  Item.Description,
  Item.Cost AS Old,
  Item.Price AS New,
  [Price] - [Cost] AS PriceDifference,
  Item.Onhand,
  [onhand] * [Cost] AS OldValue,
  ([onhand] * [Price]) AS NewValue,
  IIf(
    ([Price] > [Cost]),
    ([Price] - [Cost])* [onhand],
    0
  ) AS Markup,
  IIf(
    ([Price] < [Cost]),
    ([Price] - [Cost])* [onhand],
    0
  ) AS MarkDown
FROM
  Item
WHERE
  (
    (
      (Item.Cost)<> [Price]
    )
    AND (
      (Item.SalesDepartmentID)= 1
    )
  )
ORDER BY
  Item.ID;
