SELECT
  [vw_WebLogErrorCode-1].LicenseeName,
  [vw_WebLogErrorCode-1].WebOrderNotes,
  OrdHead.InvoiceNo,
  OrdHead.InvDate,
  OrderLine.ItemID,
  vwItemsWithCurrentPricing.Description,
  OrderLine.QuantityOrdered,
  OrderLine.Price
FROM
  (
    (
      [vw_WebLogErrorCode-1]
      INNER JOIN OrdHead ON [vw_WebLogErrorCode-1].WebOrderID = OrdHead.InvoiceNo
    )
    INNER JOIN OrderLine ON OrdHead.ID = OrderLine.OrdheadID
  )
  INNER JOIN vwItemsWithCurrentPricing ON OrderLine.ItemID = vwItemsWithCurrentPricing.ID;
