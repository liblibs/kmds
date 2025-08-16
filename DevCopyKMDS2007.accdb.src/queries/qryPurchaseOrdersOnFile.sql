SELECT
  PurchaseOrderLine.VendorID,
  Vendor.Name,
  PurchaseOrderLine.PurchaseOrderID,
  Min(PurchaseOrderLine.SetupDate) AS DateSetup
FROM
  Vendor
  INNER JOIN PurchaseOrderLine ON Vendor.ID = PurchaseOrderLine.VendorID
GROUP BY
  PurchaseOrderLine.VendorID,
  Vendor.Name,
  PurchaseOrderLine.PurchaseOrderID;
