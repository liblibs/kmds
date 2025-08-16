PARAMETERS parVendorID Long;
SELECT
  qryPurchaseOrderItemInfo.VendorID,
  qryPurchaseOrderItemInfo.Name,
  qryPurchaseOrderItemInfo.ItemID,
  qryPurchaseOrderItemInfo.SalesDepartmentID,
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
  qryPurchaseOrderItemInfo.ExtCost,
  IIf(
    (
      ([soq] \[UnitsPerCase])* [UnitsPerCase]
    )= [soq],
    0,
    -1
  ) AS Repack
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
  )
ORDER BY
  qryPurchaseOrderItemInfo.SalesDepartmentID,
  qryPurchaseOrderItemInfo.Description;
