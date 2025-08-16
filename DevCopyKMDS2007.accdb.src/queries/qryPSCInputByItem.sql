SELECT
  vwItemsWithCurrentPricing.ID,
  vwItemsWithCurrentPricing.Description,
  Sum(PSCInput.Quantity) AS SumOfQuantity
FROM
  (
    PSCInput
    INNER JOIN UPCItem ON PSCInput.Code = UPCItem.ID
  )
  INNER JOIN vwItemsWithCurrentPricing ON UPCItem.ItemID = vwItemsWithCurrentPricing.ID
GROUP BY
  vwItemsWithCurrentPricing.ID,
  vwItemsWithCurrentPricing.Description;
