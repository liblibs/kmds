SELECT
  InvoLine.Code,
  TranTypes.Description,
  InvoLine.InputCode,
  Cashier.Name AS Cashier,
  InvoLine.CashSysID AS [Reg#],
  InvoLine.InvoHeadID AS [Inv#],
  InvoLine.SetupDate,
  InvoLine.Quantity,
  InvoLine.Price,
  InvoLine.InvoHeadID,
  InvoLine.ID,
  InvoLine.CashSysID,
  InvoLine.TransType
FROM
  (
    InvoLine
    INNER JOIN TranTypes ON InvoLine.TransType = TranTypes.Code
  )
  LEFT JOIN Cashier ON InvoLine.CashierID = Cashier.ID
WHERE
  (
    (
      (InvoLine.Code)= Forms!Item!PurchaseOrdersOnFileSubItemMovement.Form!Code
    )
    And (
      (TranTypes.Description)= Forms!Item!PurchaseOrdersOnFileSubItemMovement.Form!TransType
    )
    And (
      (InvoLine.SetupDate) Between CDate(
        Forms!Item!PurchaseOrdersOnFileSubItemMovement.Form!txtDateBegin
      )
      And CDate(
        Forms!Item!PurchaseOrdersOnFileSubItemMovement.Form!txtDateEnd
      )
    )
  )
ORDER BY
  InvoLine.SetupDate;
