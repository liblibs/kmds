SELECT
  Item.ID,
  Item.Description,
  InvoLine.SetupDate,
  TranTypes.Description AS Trans,
  InvoLine.Price,
  InvoLine.InputCode,
  IIf(
    "Stock Receipts" = [Trans], [Quantity],
    0
  ) AS Purchases,
  IIf(
    " Sales - Counter" = [Trans], [Quantity],
    0
  ) AS [Counter],
  IIf(
    " Sales - Clubs" = [Trans], [Quantity],
    0
  ) AS Clubs,
  InvoLine.Quantity,
  IIf(
    "Transfers" = [Trans], [Quantity],
    0
  ) AS Transfers,
  IIf(
    "Variance Adjustment" = [Trans], [Quantity],
    0
  ) AS Adjustments,
  Cashier.Name AS Emp,
  InvoLine.TransType
FROM
  (
    (
      InvoLine
      INNER JOIN Item ON InvoLine.Code = Item.ID
    )
    INNER JOIN TranTypes ON InvoLine.TransType = TranTypes.Code
  )
  LEFT JOIN Cashier ON InvoLine.CashierID = Cashier.ID
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
  InvoLine.SetupDate;
