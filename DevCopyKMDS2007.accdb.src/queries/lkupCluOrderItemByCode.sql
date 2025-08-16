SELECT
  vwUPCItemsWithCurrentPricing.UPCCode AS CodeOrUPC,
  vwUPCItemsWithCurrentPricing.Description,
  vwUPCItemsWithCurrentPricing.PkgSize,
  vwUPCItemsWithCurrentPricing.OnShelf,
  vwUPCItemsWithCurrentPricing.Cost,
  vwUPCItemsWithCurrentPricing.CurPrice,
  IIf([SPrice] <> 0, "*", " ") AS Star,
  vwUPCItemsWithCurrentPricing.StockNumber,
  vwUPCItemsWithCurrentPricing.UnitsPerCase,
  vwUPCItemsWithCurrentPricing.Reserved,
  vwUPCItemsWithCurrentPricing.Ordered,
  vwUPCItemsWithCurrentPricing.SalesDepartmentID,
  vwUPCItemsWithCurrentPricing.ID,
  vwUPCItemsWithCurrentPricing.BuyDown,
  vwUPCItemsWithCurrentPricing.ItemVendorID,
  vwUPCItemsWithCurrentPricing.DaysSupplyGroupID,
  vwUPCItemsWithCurrentPricing.AverageDailySales,
  vwUPCItemsWithCurrentPricing.DaysSupply,
  IIf(
    [AverageDailySales] = 0, 0, [OnShelf] / [AverageDailySales]
  ) AS CurDaysSupply
FROM
  vwUPCItemsWithCurrentPricing
WHERE
  (
    (
      (
        vwUPCItemsWithCurrentPricing.LenID
      )> 2
    )
  )
ORDER BY
  vwUPCItemsWithCurrentPricing.LenID,
  vwUPCItemsWithCurrentPricing.UPCCode;
