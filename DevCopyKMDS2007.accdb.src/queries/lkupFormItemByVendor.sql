SELECT
  vwItemsWithCurrentPricing.VendorName,
  vwItemsWithCurrentPricing.StockNumber,
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
  vwItemsWithCurrentPricing.VendorName,
  vwItemsWithCurrentPricing.StockNumber,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.PkgSize;
