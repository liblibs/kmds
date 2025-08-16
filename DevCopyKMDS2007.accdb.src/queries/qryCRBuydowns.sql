PARAMETERS parInvoiceNo Long;
SELECT
  0 - InvoLine.Quantity AS BuyDownQty,
  0 - [Price] AS BuydownPrice,
  InvoLine.SalesDept AS BuyDownItemID
FROM
  InvoLine
  INNER JOIN BuyDownItem ON InvoLine.SalesDept = BuyDownItem.ItemsCode
WHERE
  (
    (
      (InvoLine.InvoHeadID)= [parInvoiceNo]
    )
    AND (
      (InvoLine.InputCode) Like "Buydown%"
    )
  )
ORDER BY
  InvoLine.InvoHeadID;
