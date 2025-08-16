INSERT INTO PurchaseOrderLine (
  PurchaseOrderID, VendorID, ItemID,
  QuantityOrdered, QuantityShipped,
  LineID, Cost
)
SELECT
  Forms!Vendors!txtNewPO AS PO,
  Forms!Vendors!ID AS vid,
  InvoLine.Code,
  InvoLine.Quantity AS Q,
  InvoLine.Quantity AS q1,
  InvoLine.ID,
  [Cost] / [Quantity] AS UCost
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.Quantity)<> 0
    )
    And (
      (InvoLine.TransType)= "P"
    )
    And (
      (InvoLine.InvoHeadID)= Forms!Vendors!txtInvoiceHeadID
    )
  );
