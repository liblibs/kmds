SELECT
  vwItemsWithCurrentPricing.PriceGroupDescr,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.PkgSize,
  vwItemsWithCurrentPricing.ID,
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
  vwItemsWithCurrentPricing.PriceGroupDescr,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.PkgSize;
