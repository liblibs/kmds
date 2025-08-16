SELECT
  vwItemsWithCurrentPricing.CategoryDescr,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.PkgSize,
  vwItemsWithCurrentPricing.ID,
  vwItemsWithCurrentPricing.OnShelf,
  vwItemsWithCurrentPricing.CurPrice,
  vwItemsWithCurrentPricing.Cost
FROM
  vwItemsWithCurrentPricing
ORDER BY
  vwItemsWithCurrentPricing.CategoryDescr,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.PkgSize;
