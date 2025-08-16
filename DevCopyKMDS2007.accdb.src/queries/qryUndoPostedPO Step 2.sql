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
  InvoLine.Quantity,
  InvoLine.Quantity,
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
