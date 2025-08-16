SELECT
  vwUPCItemsWithCurrentPricing.UPCCode,
  vwUPCItemsWithCurrentPricing.Description,
  vwUPCItemsWithCurrentPricing.PkgSize,
  vwUPCItemsWithCurrentPricing.ID,
  vwUPCItemsWithCurrentPricing.OnShelf,
  vwUPCItemsWithCurrentPricing.CurPrice,
  vwUPCItemsWithCurrentPricing.Cost,
  Format(
    ([CurPrice] - [cost])/ [CurPrice],
    "Percent"
  ) AS GP
FROM
  vwUPCItemsWithCurrentPricing
WHERE
  (
    (
      (
        vwUPCItemsWithCurrentPricing.CurPrice
      )= 0
    )
  )
ORDER BY
  vwUPCItemsWithCurrentPricing.LenID,
  vwUPCItemsWithCurrentPricing.UPCCode;
