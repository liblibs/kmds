SELECT
  vwUPCItemsWithCurrentPricing.UPCCode AS CodeOrUPC,
  vwUPCItemsWithCurrentPricing.Description,
  vwUPCItemsWithCurrentPricing.PkgSize,
  vwUPCItemsWithCurrentPricing.OnShelf,
  vwUPCItemsWithCurrentPricing.Cost,
  vwUPCItemsWithCurrentPricing.CurPrice,
  IIf([CurPrice] <> [Price], "*", " ") AS Star,
  vwUPCItemsWithCurrentPricing.StockNumber,
  vwUPCItemsWithCurrentPricing.UnitsPerCase,
  vwUPCItemsWithCurrentPricing.Reserved,
  vwUPCItemsWithCurrentPricing.Ordered,
  vwUPCItemsWithCurrentPricing.SalesDepartmentID,
  vwUPCItemsWithCurrentPricing.ID,
  vwUPCItemsWithCurrentPricing.BuyDown
FROM
  vwUPCItemsWithCurrentPricing
WHERE
  (
    (
      (
        vwUPCItemsWithCurrentPricing.LenID
      )> 3
    )
  )
ORDER BY
  vwUPCItemsWithCurrentPricing.LenID,
  vwUPCItemsWithCurrentPricing.UPCCode;
