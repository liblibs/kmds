SELECT
  Club.ID AS ClubID,
  Club.Name AS ClubName,
  InvoLine.InvoHeadID,
  Sum(InvoLine.Quantity) AS TotalQty,
  InvoLine.SetupDate AS PostDate,
  Sum(InvoLine.Cost) AS TotalCost,
  InvoLine.CashSysID,
  Sum(
    IIf([TransType] = 'C', [Price], 0)
  ) AS Retail,
  Sum(
    IIf([TransType] = 'D', [Price], 0)
  ) AS Discount,
  Sum(InvoLine.Price) AS TotalPrice,
  InvoLine.SalesGroup AS InvoiceNumber
FROM
  InvoLine
  INNER JOIN Club ON InvoLine.SalesDept = Club.ID
WHERE
  (
    (
      (InvoLine.TransType)= "C"
      Or (InvoLine.TransType)= "D"
    )
  )
GROUP BY
  Club.ID,
  Club.Name,
  InvoLine.InvoHeadID,
  InvoLine.SetupDate,
  InvoLine.CashSysID,
  InvoLine.SalesGroup
HAVING
  (
    (
      (Club.ID)= [Forms]![Clubs]![ID]
    )
  )
ORDER BY
  InvoLine.SetupDate DESC;
