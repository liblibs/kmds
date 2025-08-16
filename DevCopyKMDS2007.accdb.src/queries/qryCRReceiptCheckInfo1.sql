SELECT
  InvoLine.InvoHeadID,
  InvoLine.CashSysID,
  Tender.GetCheckNumber,
  InvoLine.SalesGroup AS CheckNumber,
  InvoLine.SalesDept AS CheckCashingID,
  [FirstName] & " " & [LastName] AS Customer,
  0 - [Price] AS CheckAmount,
  Tender.GetCheckCashing
FROM
  (
    InvoLine
    INNER JOIN Tender ON InvoLine.Code = Tender.ID
  )
  LEFT JOIN CheckCashing ON InvoLine.SalesDept = CheckCashing.ID
WHERE
  (
    (
      (InvoLine.InvoHeadID)= [parInvoiceNo]
    )
    AND (
      (InvoLine.CashSysID)= [parRegisterNumber]
    )
    AND (
      (Tender.GetCheckNumber)<> 0
    )
  )
  OR (
    (
      (InvoLine.InvoHeadID)= [parInvoiceNo]
    )
    AND (
      (InvoLine.CashSysID)= [parRegisterNumber]
    )
    AND (
      (Tender.GetCheckCashing)<> 0
    )
  );
