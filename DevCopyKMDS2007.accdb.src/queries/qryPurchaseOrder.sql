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
  Item.Description,
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
  Item.UnitsPerCase,
  [QuantityShipped] * PurchaseOrderLine.Cost AS ExtCost,
  [QuantityShipped] * PurchaseOrderLine.Price AS ExtPrice,
  Item.StockNumber,
  qryUPCItemMax.UPCCodeMax,
  Vendor.Phone2 AS FAX,
  Vendor.PoAtCost
FROM
  (
    Vendor
    INNER JOIN (
      PurchaseOrder
      INNER JOIN (
        (
          Item
          LEFT JOIN [Size] ON Item.SizeID = Size.ID
        )
        INNER JOIN PurchaseOrderLine ON Item.ID = PurchaseOrderLine.ItemID
      ) ON PurchaseOrder.ID = PurchaseOrderLine.PurchaseOrderID
    ) ON Vendor.ID = PurchaseOrderLine.VendorID
  )
  LEFT JOIN qryUPCItemMax ON Item.ID = qryUPCItemMax.ItemID
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
