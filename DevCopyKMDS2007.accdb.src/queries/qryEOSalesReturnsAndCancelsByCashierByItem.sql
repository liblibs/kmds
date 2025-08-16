SELECT
  Cashier.Name,
  vwItemsWithCurrentPricing.ID,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.OnShelf,
  InvoLine5.SetupDate,
  InvoLine5.CashSysID,
  InvoLine5.InvoHeadID,
  InvoLine5.InputCode,
  InvoLine5.Quantity,
  InvoLine5.Price
FROM
  (
    InvoLine5
    INNER JOIN vwItemsWithCurrentPricing ON InvoLine5.Code = vwItemsWithCurrentPricing.ID
  )
  INNER JOIN Cashier ON InvoLine5.CashierID = Cashier.ID
WHERE
  (
    (
      (InvoLine5.Quantity)< 0
    )
    AND (
      (InvoLine5.TransType)= "S"
    )
  )
ORDER BY
  Cashier.Name,
  vwItemsWithCurrentPricing.ID;
