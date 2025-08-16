SELECT
  vwItemsWithCurrentPricing.SalesDptDescr,
  vwItemsWithCurrentPricing.CategoryDescr AS Category,
  Item.ID,
  vwItemsWithCurrentPricing.Description AS Description,
  Size.Description AS [Size],
  Item.RestockAmount,
  [RestockAmount] * [vwItemsWithCurrentPricing].[Price] AS Dollars,
  vwItemsWithCurrentPricing.OnShelf,
  Item.Location
FROM
  vwItemsWithCurrentPricing
  INNER JOIN (
    Item
    INNER JOIN [Size] ON Item.SizeID = Size.ID
  ) ON vwItemsWithCurrentPricing.ID = Item.ID
WHERE
  (
    (
      (Item.RestockAmount)<> 0
    )
    AND (
      (Item.ID) Is Not Null
    )
  )
ORDER BY
  vwItemsWithCurrentPricing.CategoryDescr,
  Size.Description,
  Item.ID;
