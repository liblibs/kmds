SELECT
  PSCInput.SalesDepartmentID,
  vwUPCItemsWithCurrentPricing.ID,
  Sum(PSCInput.Quantity) AS QuantityShipped,
  vwUPCItemsWithCurrentPricing.CurPrice,
  vwUPCItemsWithCurrentPricing.Cost,
  vwUPCItemsWithCurrentPricing.BuyDown,
  vwUPCItemsWithCurrentPricing.OnShelf
FROM
  PSCInput
  INNER JOIN vwUPCItemsWithCurrentPricing ON PSCInput.Code = vwUPCItemsWithCurrentPricing.UPCCode
GROUP BY
  PSCInput.SalesDepartmentID,
  vwUPCItemsWithCurrentPricing.ID,
  vwUPCItemsWithCurrentPricing.CurPrice,
  vwUPCItemsWithCurrentPricing.Cost,
  vwUPCItemsWithCurrentPricing.BuyDown,
  vwUPCItemsWithCurrentPricing.OnShelf
ORDER BY
  PSCInput.SalesDepartmentID,
  vwUPCItemsWithCurrentPricing.ID;
