SELECT
  BuyDown.ID,
  BuyDown.RepsName,
  BuyDown.ReferenceNo,
  BuyDown.Club,
  BuyDown.Order,
  BuyDown.InvoiceNo,
  BuyDown.BeginDate,
  BuyDown.EndDate,
  BuyDown.BuydownDollars,
  BuyDown.DollarsToDate,
  BuyDown.BuydownQty,
  BuyDown.Quantity,
  BuyDown.LastSoldDate,
  BuyDown.Status,
  BuyDown.PaidFlag,
  BuyDown.ItemsID,
  BuyDown.SetupDate,
  BuyDown.ChangeDate,
  BuyDownItem.ItemsCode,
  Item.Description,
  BuyDownItem.QuantitySold,
  BuyDownItem.DiscountDollars
FROM
  (
    BuyDown
    INNER JOIN BuyDownItem ON BuyDown.ID = BuyDownItem.BuyDownID
  )
  INNER JOIN Item ON BuyDownItem.ItemsCode = Item.ID
WHERE
  (
    (
      (BuyDown.ID)= Forms!BuyDowns!ID
    )
    And (
      (BuyDownItem.QuantitySold)<> 0
    )
  );
