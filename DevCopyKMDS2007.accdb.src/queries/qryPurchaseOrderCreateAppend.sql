SELECT
  qryPurchaseOrderItemInfo.VendorID,
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
    AND (
      (qryPurchaseOrderItemInfo.Avail)<= [OrderAt]
    )
    AND (
      (
        qryPurchaseOrderItemInfo.OrderTo
      )> 0
    )
    AND (
      (qryPurchaseOrderItemInfo.SOQ)> 0
    )
  );
