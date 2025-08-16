SELECT
  qryInvoLineReceiptsWithItemsSelected.DateOfSale,
  qryInvoLineReceiptsWithItemsSelected.CashSysID,
  qryInvoLineReceiptsWithItemsSelected.InvoHeadID,
  InvoLine.ID,
  IIf(
    [TransType] = "T", "Tender", "Sale"
  ) AS Trans,
  InvoLine.CashierID,
  IIf(
    IsNull(Cashier.Name),
    "Deleted",
    Cashier.Name
  ) AS Cashier,
  InvoLine.Code,
  Item.Description,
  InvoLine.Quantity,
  InvoLine.Price,
  InvoLine.InputCode
FROM
  (
    (
      qryInvoLineReceiptsWithItemsSelected
      INNER JOIN InvoLine ON (
        qryInvoLineReceiptsWithItemsSelected.InvoHeadID = InvoLine.InvoHeadID
      )
      AND (
        qryInvoLineReceiptsWithItemsSelected.CashSysID = InvoLine.CashSysID
      )
    )
    LEFT JOIN Cashier ON InvoLine.CashierID = Cashier.ID
  )
  LEFT JOIN Item ON InvoLine.Code = Item.ID
ORDER BY
  qryInvoLineReceiptsWithItemsSelected.DateOfSale,
  InvoLine.ID;
