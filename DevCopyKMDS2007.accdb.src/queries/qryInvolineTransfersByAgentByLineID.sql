SELECT
  InvoLine.SalesDept AS [Store Number],
  InvoLine.SetupDate AS PostDate,
  Agents.LastName AS StoreName,
  InvoLine.SalesGroup AS TransferNumber,
  InvoLine.InvoHeadID,
  InvoLine.ID,
  InvoLine.Code,
  Item.Description,
  InvoLine.Quantity,
  InvoLine.Cost,
  InvoLine.Price
FROM
  (
    InvoLine
    INNER JOIN Agents ON InvoLine.SalesDept = Agents.AgentNumber
  )
  INNER JOIN Item ON InvoLine.Code = Item.ID
WHERE
  (
    (
      (InvoLine.TransType)= "X"
    )
  )
ORDER BY
  InvoLine.SalesDept,
  InvoLine.SetupDate DESC,
  InvoLine.ID;
