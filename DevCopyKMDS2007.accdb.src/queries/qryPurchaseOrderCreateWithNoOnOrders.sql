PARAMETERS parVendorID Long;
SELECT
  qryPurchaseOrderItemInfoWithNoOnOrder.VendorID,
  qryPurchaseOrderItemInfoWithNoOnOrder.Name,
  qryPurchaseOrderItemInfoWithNoOnOrder.ItemID,
  qryPurchaseOrderItemInfoWithNoOnOrder.SalesDepartmentID,
  qryPurchaseOrderItemInfoWithNoOnOrder.Description,
  qryPurchaseOrderItemInfoWithNoOnOrder.SizeID,
  qryPurchaseOrderItemInfoWithNoOnOrder.Onhand,
  qryPurchaseOrderItemInfoWithNoOnOrder.Ordered,
  qryPurchaseOrderItemInfoWithNoOnOrder.Reserved,
  qryPurchaseOrderItemInfoWithNoOnOrder.Avail,
  qryPurchaseOrderItemInfoWithNoOnOrder.OrderAt,
  qryPurchaseOrderItemInfoWithNoOnOrder.OrderTo,
  qryPurchaseOrderItemInfoWithNoOnOrder.SOQRaw,
  qryPurchaseOrderItemInfoWithNoOnOrder.OrderMultiple,
  qryPurchaseOrderItemInfoWithNoOnOrder.SOQ,
  qryPurchaseOrderItemInfoWithNoOnOrder.UnitsPerCase,
  qryPurchaseOrderItemInfoWithNoOnOrder.SPrice,
  qryPurchaseOrderItemInfoWithNoOnOrder.SPriceStar,
  qryPurchaseOrderItemInfoWithNoOnOrder.ExtPrice,
  qryPurchaseOrderItemInfoWithNoOnOrder.Cost,
  qryPurchaseOrderItemInfoWithNoOnOrder.ExtCost,
  IIf(
    (
      ([soq] \[UnitsPerCase])* [UnitsPerCase]
    )= [soq],
    0,
    -1
  ) AS Repack,
  qryPurchaseOrderItemInfoWithNoOnOrder.SPrice,
  qryPurchaseOrderItemInfoWithNoOnOrder.SPrice
FROM
  qryPurchaseOrderItemInfoWithNoOnOrder
WHERE
  (
    (
      (
        qryPurchaseOrderItemInfoWithNoOnOrder.VendorID
      )= [parVendorID]
    )
    AND (
      (
        qryPurchaseOrderItemInfoWithNoOnOrder.Avail
      )<= [OrderAt]
    )
    AND (
      (
        qryPurchaseOrderItemInfoWithNoOnOrder.OrderTo
      )> 0
    )
    AND (
      (
        qryPurchaseOrderItemInfoWithNoOnOrder.SOQ
      )> 0
    )
  );
