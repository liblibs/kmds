SELECT
  Cashier.Name,
  vwItemsWithCurrentPricing.ID,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.OnShelf,
  InvoLine.SetupDate,
  InvoLine.CashSysID,
  InvoLine.InvoHeadID,
  InvoLine.InputCode,
  InvoLine.Quantity,
  InvoLine.Price
FROM
  (
    InvoLine
    INNER JOIN Cashier ON InvoLine.CashierID = Cashier.ID
  )
  INNER JOIN vwItemsWithCurrentPricing ON InvoLine.Code = vwItemsWithCurrentPricing.ID
WHERE
  (
    (
      (InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
    And (
      (InvoLine.Quantity)< 0
    )
    And (
      (InvoLine.TransType)= "S"
    )
  )
ORDER BY
  Cashier.Name,
  vwItemsWithCurrentPricing.ID;
