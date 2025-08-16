SELECT
  PriceGroup.ID AS PriceGrouipID,
  PriceGroup.Description AS PriceGroupDesc,
  Item.ID AS ItemID,
  Item.Description,
  InvoLine.SetupDate,
  TranTypes.Description AS Trans,
  InvoLine.Price,
  InvoLine.InputCode,
  IIf(
    "P" = TranTypes.Code, [Quantity], 0
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
    "X" = TranTypes.Code, [Quantity], 0
  ) AS Transfers,
  IIf(
    "A" = TranTypes.Code, [Quantity], 0
  ) AS Adjustments,
  Cashier.Name AS Emp,
  TranTypes.Code AS TranType
FROM
  PriceGroup
  INNER JOIN (
    (
      (
        InvoLine
        INNER JOIN Item ON InvoLine.Code = Item.ID
      )
      INNER JOIN TranTypes ON InvoLine.TransType = TranTypes.Code
    )
    LEFT JOIN Cashier ON InvoLine.CashierID = Cashier.ID
  ) ON PriceGroup.ID = Item.PriceGroupID
WHERE
  (
    (
      (PriceGroup.Description) Between Forms!PrintInventoryReport!cboPriceGroupIDBegin
      And Forms!PrintInventoryReport!cboPriceGroupIDEnd
    )
    And (
      (InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
  )
ORDER BY
  PriceGroup.ID,
  Item.Description,
  InvoLine.SetupDate;
