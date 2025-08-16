SELECT
  InvoHead.ID AS InvoiceNumber,
  InvoHead.CashSysID,
  Cashier.Name AS CashierName,
  InvoHead.SetupDate,
  InvoLine.TransType,
  InvoLine.Code,
  InvoLine.SalesDept AS ODL,
  InvoLine.InputCode AS DOB
FROM
  (
    Cashier
    INNER JOIN (
      InvoHead
      INNER JOIN InvoLine ON (
        InvoHead.CashSysID = InvoLine.CashSysID
      )
      AND (
        InvoHead.ID = InvoLine.InvoHeadID
      )
    ) ON Cashier.ID = InvoHead.CashierID
  )
  INNER JOIN Tender ON InvoLine.Code = Tender.ID
WHERE
  (
    (
      (InvoHead.ID)= [parInvoiceNo]
    )
    AND (
      (InvoHead.CashSysID)= [parRegisterNumber]
    )
    AND (
      (InvoLine.TransType)= "X"
    )
  );
