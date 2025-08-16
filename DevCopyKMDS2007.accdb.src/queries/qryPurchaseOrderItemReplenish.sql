PARAMETERS parBeginDate DateTime,
parEndDate DateTime,
parVendorID Long,
parBeginSalesDepartmentID Long,
parEndSalesDepartmentID Long;
SELECT
  qryPurchaseOrderItemDaysSupplyADS.VendorID,
  Vendor.Name AS VendorName,
  qryPurchaseOrderItemDaysSupplyADS.ItemId,
  qryVWItemsWithCurrentPricing.Description,
  qryVWItemsWithCurrentPricing.PkgSize AS SizeDesc,
  qryPurchaseOrderItemDaysSupplyADS.Sold,
  qryPurchaseOrderItemDaysSupplyADS.SoldDays,
  qryVWItemsWithCurrentPricing.Ordered,
  qryVWItemsWithCurrentPricing.Reserved,
  [OnShelf] + [Ordered] AS Avail,
  qryPurchaseOrderItemDaysSupplyADS.Sold AS SOQRaw,
  qryVWItemsWithCurrentPricing.UnitsPerOrder AS OrderMultiple,
  CInt(
    (
      ([Sold])/ [UnitsPerOrder]
    )
  )* [UnitsPerOrder] AS SOQ,
  qryPurchaseOrderItemDaysSupplyADS.DateStart,
  qryPurchaseOrderItemDaysSupplyADS.DateEnd,
  qryVWItemsWithCurrentPricing.UnitsPerCase,
  IIf(
    [soq] Mod [UnitsPerCase] <> 0, True,
    False
  ) AS Repack,
  qryVWItemsWithCurrentPricing.CurPrice AS SPrice,
  IIf([Price] = [CurPrice], " ", "*") AS SPriceStar,
  [SOQ] * [SPrice] AS ExtPrice,
  qryVWItemsWithCurrentPricing.Cost,
  [SOQ] * [Cost] AS ExtCost
FROM
  (
    Vendor
    INNER JOIN qryPurchaseOrderItemDaysSupplyADS ON Vendor.ID = qryPurchaseOrderItemDaysSupplyADS.VendorID
  )
  INNER JOIN qryVWItemsWithCurrentPricing ON qryPurchaseOrderItemDaysSupplyADS.ItemId = qryVWItemsWithCurrentPricing.ID
WHERE
  (
    (
      (
        qryPurchaseOrderItemDaysSupplyADS.Sold
      )<> 0
    )
  )
ORDER BY
  qryVWItemsWithCurrentPricing.SalesDepartmentID,
  qryVWItemsWithCurrentPricing.Description;
