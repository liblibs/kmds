PARAMETERS parBeginDate DateTime,
parEndDate DateTime,
parVendorID Long,
parBeginSalesDepartmentID Long,
parEndSalesDepartmentID Long;
SELECT
  qryPurchaseOrderItemDaysSupplyADS.VendorID,
  Vendor.Name AS VendorName,
  qryPurchaseOrderItemDaysSupplyADS.ItemId,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.PkgSize AS SizeDesc,
  qryPurchaseOrderItemDaysSupplyADS.Sold,
  qryPurchaseOrderItemDaysSupplyADS.SoldDays,
  vwItemsWithCurrentPricing.Ordered,
  vwItemsWithCurrentPricing.Reserved,
  [OnShelf] + [Ordered] AS Avail,
  qryPurchaseOrderItemDaysSupplyADS.Sold AS SOQRaw,
  vwItemsWithCurrentPricing.UnitsPerOrder AS OrderMultiple,
  CInt(
    (
      ([Sold])/ [UnitsPerOrder]
    )
  )* [UnitsPerOrder] AS SOQ,
  qryPurchaseOrderItemDaysSupplyADS.DateStart,
  qryPurchaseOrderItemDaysSupplyADS.DateEnd,
  vwItemsWithCurrentPricing.UnitsPerCase,
  IIf(
    [soq] Mod [UnitsPerCase] <> 0, True,
    False
  ) AS Repack,
  vwItemsWithCurrentPricing.CurPrice AS SPrice,
  IIf([Price] = [CurPrice], " ", "*") AS SPriceStar,
  [SOQ] * [SPrice] AS ExtPrice,
  vwItemsWithCurrentPricing.Cost,
  [SOQ] * [Cost] AS ExtCost
FROM
  (
    Vendor
    INNER JOIN qryPurchaseOrderItemDaysSupplyADS ON Vendor.ID = qryPurchaseOrderItemDaysSupplyADS.VendorID
  )
  INNER JOIN vwItemsWithCurrentPricing ON qryPurchaseOrderItemDaysSupplyADS.ItemId = vwItemsWithCurrentPricing.ID
WHERE
  (
    (
      (
        qryPurchaseOrderItemDaysSupplyADS.Sold
      )<> 0
    )
  )
ORDER BY
  vwItemsWithCurrentPricing.SalesDepartmentID,
  vwItemsWithCurrentPricing.Description;
