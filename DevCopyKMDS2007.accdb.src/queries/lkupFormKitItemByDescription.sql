SELECT
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
WHERE
  (
    (
      (
        vwItemsWithCurrentPricing.CurPrice
      )= 0
    )
  )
ORDER BY
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.PkgSize;
