SELECT
  Item.ID,
  Item.Description,
  InvoLine.Quantity,
  InvoLine.Price,
  InvoLine.SetupDate,
  Format(
    InvoLine.SetupDate, "mm/dd/yyyy"
  ) AS BusDate
FROM
  (
    BuyDown
    INNER JOIN Item ON BuyDown.ItemsID = Item.ID
  )
  INNER JOIN InvoLine ON Item.ID = InvoLine.Code
WHERE
  (
    (
      (Item.ID) Between Forms!PrintInventoryReport!cboItemIDBegin
      And Forms!PrintInventoryReport!cboItemIDEnd
    )
    And (
      (InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
  )
ORDER BY
  Item.ID,
  InvoLine.SetupDate;
