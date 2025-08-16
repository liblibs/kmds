SELECT
  InvoLine.SalesDept AS [Store Number],
  InvoLine.SetupDate AS PostDate,
  Agents.LastName AS StoreName,
  First(InvoLine.SalesGroup) AS TransferNumber,
  InvoLine.InvoHeadID,
  Sum(InvoLine.Quantity) AS TotalQty,
  Sum(InvoLine.Price) AS TotalPrice,
  Sum(InvoLine.Cost) AS TotalCost
FROM
  InvoLine
  INNER JOIN Agents ON InvoLine.SalesDept = Agents.AgentNumber
WHERE
  (
    (
      (InvoLine.TransType)= "X"
    )
  )
GROUP BY
  InvoLine.SalesDept,
  InvoLine.SetupDate,
  Agents.LastName,
  InvoLine.InvoHeadID
ORDER BY
  InvoLine.SalesDept,
  InvoLine.SetupDate DESC;
