PARAMETERS parInvoiceNo Long,
parRegisterNumber Long;
SELECT
  InvoHead.ID AS InvoiceNumber,
  InvoHead.CashSysID,
  Cashier.Name AS CashierName,
  InvoHead.SetupDate,
  InvoLine.ID,
  InvoLine.TransType,
  InvoLine.Code,
  InvoLine.InputCode AS TDescription,
  Item.Description AS IDescription,
  InvoLine.Quantity,
  InvoLine.Price,
  InvoLine.InputCode
FROM
  (
    (
      Cashier
      INNER JOIN (
        InvoHead
        INNER JOIN InvoLine ON (
          InvoHead.ID = InvoLine.InvoHeadID
        )
        AND (
          InvoHead.CashSysID = InvoLine.CashSysID
        )
      ) ON Cashier.ID = InvoHead.CashierID
    )
    LEFT JOIN Tender ON InvoLine.Code = Tender.ID
  )
  LEFT JOIN Item ON InvoLine.Code = Item.ID
WHERE
  (
    (
      (InvoHead.ID)= [parInvoiceNo]
    )
    AND (
      (InvoHead.CashSysID)= [parRegisterNumber]
    )
  )
ORDER BY
  InvoLine.ID;
