SELECT
  BuyDown.RepsName,
  BuyDown.ReferenceNo,
  Item.Description,
  BuyDownItem.ItemsCode,
  BuyDownItem.BuyDownID,
  InvoLine.SetupDate AS InvoiceDate,
  InvoLine.CashSysID AS CashRegister,
  InvoLine.InvoHeadID AS InvoiceNumber,
  InvoLine.Quantity,
  InvoLine.Price
FROM
  BuyDown
  INNER JOIN (
    (
      BuyDownItem
      INNER JOIN Item ON BuyDownItem.ItemsCode = Item.ID
    )
    INNER JOIN InvoLine ON Item.ID = InvoLine.Code
  ) ON BuyDown.ID = BuyDownItem.BuyDownID
WHERE
  (
    (
      (BuyDownItem.BuyDownID)= [parBuydownNumber]
    )
    And (
      (InvoLine.SetupDate) Between buydown.BeginDate
      And buydown.EndDate
    )
    And (
      (InvoLine.TransType)= "S"
      Or (InvoLine.TransType)= "C"
    )
  )
ORDER BY
  Item.Description;
