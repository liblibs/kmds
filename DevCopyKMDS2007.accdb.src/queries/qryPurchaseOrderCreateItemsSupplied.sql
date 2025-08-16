PARAMETERS parVendorID Long;
SELECT
  qryPurchaseOrderItemInfo.VendorID,
  qryPurchaseOrderItemInfo.SalesDepartmentID,
  qryPurchaseOrderItemInfo.Name,
  qryPurchaseOrderItemInfo.ItemID,
  qryPurchaseOrderItemInfo.Description,
  qryPurchaseOrderItemInfo.SizeID,
  qryPurchaseOrderItemInfo.Onhand,
  qryPurchaseOrderItemInfo.Ordered,
  qryPurchaseOrderItemInfo.Reserved,
  qryPurchaseOrderItemInfo.Avail,
  qryPurchaseOrderItemInfo.OrderAt,
  qryPurchaseOrderItemInfo.OrderTo,
  qryPurchaseOrderItemInfo.SOQRaw,
  qryPurchaseOrderItemInfo.OrderMultiple,
  qryPurchaseOrderItemInfo.SOQ,
  qryPurchaseOrderItemInfo.UnitsPerCase,
  qryPurchaseOrderItemInfo.SPrice,
  qryPurchaseOrderItemInfo.SPriceStar,
  qryPurchaseOrderItemInfo.ExtPrice,
  qryPurchaseOrderItemInfo.Cost,
  qryPurchaseOrderItemInfo.ExtCost
FROM
  qryPurchaseOrderItemInfo
WHERE
  (
    (
      (
        qryPurchaseOrderItemInfo.VendorID
      )= [parVendorID]
    )
  )
ORDER BY
  qryPurchaseOrderItemInfo.SalesDepartmentID,
  qryPurchaseOrderItemInfo.Description;
