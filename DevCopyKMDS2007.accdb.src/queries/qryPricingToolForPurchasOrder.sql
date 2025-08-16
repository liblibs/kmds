SELECT
  PricingWorkTable.ID,
  vwItemsWithCurrentPricing.Cost,
  vwItemsWithCurrentPricing.Price,
  vwItemsWithCurrentPricing.UnitsPerCase
FROM
  vwItemsWithCurrentPricing
  INNER JOIN PricingWorkTable ON vwItemsWithCurrentPricing.ID = PricingWorkTable.ID
WHERE
  (
    (
      (
        vwItemsWithCurrentPricing.Price
      )<> [PricingWorkTable].[Price]
    )
  )
ORDER BY
  PricingWorkTable.ID;
