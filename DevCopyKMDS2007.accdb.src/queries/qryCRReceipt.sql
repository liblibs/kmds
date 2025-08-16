PARAMETERS parInvoiceNo Long,
parRegisterNumber Long;
SELECT
  InvoLine.InvoHeadID AS InvoiceNumber,
  InvoLine.CashSysID,
  Cashier.Name AS CashierName,
  InvoLine.ID,
  InvoLine.TransType,
  InvoLine.Code,
  InvoLine.InputCode AS TDescription,
  Item.Description AS IDescription,
  InvoLine.Quantity,
  InvoLine.Price,
  InvoLine.InputCode,
  InvoLine.SetupDate
FROM
  Cashier
  INNER JOIN (
    (
      InvoLine
      LEFT JOIN Tender ON InvoLine.Code = Tender.ID
    )
    LEFT JOIN Item ON InvoLine.Code = Item.ID
  ) ON Cashier.ID = InvoLine.CashierID
WHERE
  (
    (
      (InvoLine.InvoHeadID)= [parInvoiceNo]
    )
    AND (
      (InvoLine.CashSysID)= [parRegisterNumber]
    )
  )
ORDER BY
  InvoLine.InvoHeadID,
  InvoLine.ID,
  InvoLine.TransType;
