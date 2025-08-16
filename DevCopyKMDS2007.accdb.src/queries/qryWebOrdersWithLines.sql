SELECT
  [vw_WebLogErrorCode-1].WebOrderID,
  [vw_WebLogErrorCode-1].LicenseeName,
  OrdHead.InvoiceNo,
  OrderLine.ItemID,
  vwItemsWithCurrentPricing.Description,
  OrderLine.QuantityOrdered,
  [Orderline].[Price] * [QuantityOrdered] AS ExtPrice,
  OrderLine.QuantityShipped
FROM
  (
    (
      [vw_WebLogErrorCode-1]
      INNER JOIN OrdHead ON [vw_WebLogErrorCode-1].WebOrderID = OrdHead.InvoiceNo
    )
    INNER JOIN OrderLine ON OrdHead.ID = OrderLine.OrdheadID
  )
  INNER JOIN vwItemsWithCurrentPricing ON OrderLine.ItemID = vwItemsWithCurrentPricing.ID;
