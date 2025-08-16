SELECT
  InvoLine.TransType,
  InvoLine.InvoHeadID,
  InvoLine.CashSysID,
  InvoLine.SalesDept AS ClubID,
  InvoLine.SalesGroup AS InvoiceNumber,
  InvoLine.SalesGroup AS [Order],
  InvoLine.SetupDate,
  Club.Name,
  InvoLine.Quantity,
  InvoLine.Price
FROM
  InvoLine
  INNER JOIN Club ON InvoLine.SalesDept = Club.ID
WHERE
  (
    (
      (InvoLine.TransType)= "D"
    )
    AND (
      (InvoLine.Quantity)=-1
    )
  )
ORDER BY
  InvoLine.InvoHeadID;
