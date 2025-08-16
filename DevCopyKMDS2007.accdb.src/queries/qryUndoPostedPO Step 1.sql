PARAMETERS VendorNumber Long,
NewPO Long,
InvoiceHeadID Long; INSERT INTO PurchaseOrderLine (
  PurchaseOrderID, VendorID, ItemID,
  QuantityOrdered, QuantityShipped,
  LineID
)
SELECT
  [NewPO] AS PO,
  [VendorNumber] AS vid,
  InvoLine.Code,
  0 - [Quantity] AS Q,
  0 - [Quantity] AS q1,
  InvoLine.ID
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.TransType)= "P"
    )
    AND (
      (InvoLine.InvoHeadID)= [InvoiceHeadID]
    )
  );
