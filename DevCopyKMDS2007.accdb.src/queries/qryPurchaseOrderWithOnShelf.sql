SELECT
  PurchaseOrder.ID,
  PurchaseOrder.VendorID,
  PurchaseOrder.OrderCount,
  PurchaseOrder.OrderDate,
  Vendor.Name AS VendorName,
  Vendor.Contact,
  Vendor.Address1,
  Vendor.Address2,
  [City] & ", " & [St] & "  " & [Zip] AS CitySTZ,
  Vendor.Phone1,
  Vendor.PoAtCost,
  Vendor.PrintUpcOnPo,
  PurchaseOrderLine.LineID,
  PurchaseOrderLine.ItemID,
  vwItemsWithCurrentPricing.Description,
  PurchaseOrderLine.QuantityShipped,
  PurchaseOrderLine.Cost,
  PurchaseOrderLine.Price,
  PurchaseOrderLine.SetupDate,
  PurchaseOrderLine.Repack,
  Size.Description AS SizeDesc,
  IIf(
    [UnitsPerCase] > 1, [QuantityShipped] \[UnitsPerCase],
    0
  ) AS Cases,
  IIf(
    [UnitsPerCase] > 1,
    [QuantityShipped] -(
      (
        [QuantityShipped] \[UnitsPerCase]
      )* [UnitsPerCase]
    ),
    [QuantityShipped]
  ) AS Singles,
  vwItemsWithCurrentPricing.UnitsPerCase,
  [QuantityShipped] * PurchaseOrderLine.Cost AS ExtCost,
  [QuantityShipped] * PurchaseOrderLine.Price AS ExtPrice,
  vwItemsWithCurrentPricing.StockNumber,
  qryUPCItemMax.UPCCodeMax,
  Vendor.Phone2 AS FAX,
  Vendor.PoAtCost,
  vwItemsWithCurrentPricing.OnShelf,
  vwItemsWithCurrentPricing.DaysSupply,
  IIf(
    [AverageDailySales] > 0
    And [Onhand] > 0,
    [Onhand] / [AverageDailySales], 0
  ) AS DaysLeft,
  vwItemsWithCurrentPricing.AverageDailySales
FROM
  (
    (
      (
        Vendor
        INNER JOIN (
          PurchaseOrder
          INNER JOIN PurchaseOrderLine ON PurchaseOrder.ID = PurchaseOrderLine.PurchaseOrderID
        ) ON Vendor.ID = PurchaseOrderLine.VendorID
      )
      INNER JOIN vwItemsWithCurrentPricing ON PurchaseOrderLine.ItemID = vwItemsWithCurrentPricing.ID
    )
    LEFT JOIN qryUPCItemMax ON vwItemsWithCurrentPricing.ID = qryUPCItemMax.ItemID
  )
  LEFT JOIN [Size] ON vwItemsWithCurrentPricing.SizeID = Size.ID
WHERE
  (
    (
      (PurchaseOrder.ID)= Forms!PurchaseOrderChoices!txtPurchaseOrderID
    )
    And (
      (
        PurchaseOrderLine.QuantityShipped
      )<> 0
    )
  );
