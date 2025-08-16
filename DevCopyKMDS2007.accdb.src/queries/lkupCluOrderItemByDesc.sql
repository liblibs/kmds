SELECT
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.ID,
  vwItemsWithCurrentPricing.PkgSize,
  vwItemsWithCurrentPricing.OnShelf,
  vwItemsWithCurrentPricing.CurPrice,
  IIf([SPrice] <> 0, "*", " ") AS Star,
  vwItemsWithCurrentPricing.SalesDepartmentID,
  vwItemsWithCurrentPricing.Cost,
  vwItemsWithCurrentPricing.StockNumber,
  vwItemsWithCurrentPricing.UnitsPerCase,
  vwItemsWithCurrentPricing.ItemVendorID,
  IIf(
    [AverageDailySales] = 0, 0, [OnShelf] / [AverageDailySales]
  ) AS CurDaysSupply
FROM
  vwItemsWithCurrentPricing
ORDER BY
  vwItemsWithCurrentPricing.Description;
