SELECT
  First(
    vwUPCItemsWithCurrentPricing.UPCCode
  ) AS FirstOfUPCCode,
  vwUPCItemsWithCurrentPricing.ID,
  vwUPCItemsWithCurrentPricing.Description,
  vwUPCItemsWithCurrentPricing.Onhand INTO aPalm
FROM
  vwUPCItemsWithCurrentPricing
WHERE
  (
    (
      (
        vwUPCItemsWithCurrentPricing.LenID
      )> 7
    )
  )
GROUP BY
  vwUPCItemsWithCurrentPricing.ID,
  vwUPCItemsWithCurrentPricing.Description,
  vwUPCItemsWithCurrentPricing.Onhand
HAVING
  (
    (
      (
        vwUPCItemsWithCurrentPricing.ID
      ) Between [BeginCode]
      And [EndCode]
    )
  )
ORDER BY
  First(
    vwUPCItemsWithCurrentPricing.UPCCode
  );
