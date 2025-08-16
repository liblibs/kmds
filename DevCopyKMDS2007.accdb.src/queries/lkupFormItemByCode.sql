SELECT
  vwItemsWithCurrentPricing.ID,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.PkgSize,
  vwItemsWithCurrentPricing.OnShelf,
  vwItemsWithCurrentPricing.CurPrice,
  vwItemsWithCurrentPricing.Cost,
  Format(
    ([CurPrice] - [cost])/ [CurPrice],
    "Percent"
  ) AS GP
FROM
  vwItemsWithCurrentPricing
ORDER BY
  vwItemsWithCurrentPricing.ID;
