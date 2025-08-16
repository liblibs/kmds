SELECT
  vwItemsWithCurrentPricing.ItemVendorID AS VendorID,
  vwItemsWithCurrentPricing.VendorName AS Name,
  vwItemsWithCurrentPricing.ID AS ItemID,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.SizeID,
  vwItemsWithCurrentPricing.Onhand,
  vwItemsWithCurrentPricing.Ordered,
  vwItemsWithCurrentPricing.Reserved,
  [OnShelf] + [Ordered] AS Avail,
  vwItemsWithCurrentPricing.AverageDailySales,
  vwItemsWithCurrentPricing.DaysSupply,
  CLng(
    IIf(
      IsNull([AverageDailySales]),
      0,
      [AverageDailySales]
    )* [DaysSupply]
  ) AS OrderMax,
  ([OrderMax] - [Avail] + 0.01) AS SOQRaw,
  IIf(
    [UnitsPerOrder] < 1, 1, [UnitsPerOrder]
  ) AS OrderMultiple,
  (
    CInt([SOQRaw] / [OrderMultiple])* [OrderMultiple]
  ) AS SOQ,
  vwItemsWithCurrentPricing.UnitsPerCase,
  vwItemsWithCurrentPricing.Price,
  vwItemsWithCurrentPricing.Sprice,
  IIf(
    ([SPrice] = 0),
    " ",
    "*"
  ) AS SPriceStar,
  [SOQ] * [SPrice] AS ExtPrice,
  vwItemsWithCurrentPricing.Cost,
  [SOQ] * [Cost] AS ExtCost,
  vwItemsWithCurrentPricing.SalesDepartmentID,
  IIf(
    [soq] Mod [UnitsPerCase] <> 0, True,
    False
  ) AS Repack,
  vwItemsWithCurrentPricing.CategoryDescr INTO PurchaseOrderItemWADS
FROM
  vwItemsWithCurrentPricing
WHERE
  (
    (
      (
        vwItemsWithCurrentPricing.ItemVendorID
      )= [parVendorID]
    )
    AND (
      (
        vwItemsWithCurrentPricing.CategoryDescr
      )<> "OBSOLETE"
    )
  )
ORDER BY
  vwItemsWithCurrentPricing.ID;
