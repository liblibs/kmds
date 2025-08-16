SELECT
  InvoLine.SalesGroup AS VendorNo,
  InvoLine.SalesDept AS PONumber,
  InvoLine.InvoHeadID,
  InvoLine.InputCode,
  Sum(InvoLine.Quantity) AS TotalQty,
  Sum(InvoLine.Price) AS TotalPrice,
  InvoLine.SetupDate AS PostDate,
  Sum(InvoLine.Cost) AS TotalCost,
  InvoLine.CashSysID
FROM
  InvoLine
GROUP BY
  InvoLine.SalesGroup,
  InvoLine.SalesDept,
  InvoLine.InvoHeadID,
  InvoLine.InputCode,
  InvoLine.SetupDate,
  InvoLine.CashSysID,
  InvoLine.TransType
HAVING
  (
    (
      (InvoLine.TransType)= "P"
    )
  )
ORDER BY
  InvoLine.SetupDate DESC;
