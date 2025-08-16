SELECT
  InvoLine.CashierID,
  Cashier.Name AS CashierName,
  InvoLine.SetupDate,
  InvoLine.CashSysID AS RegisterNumber,
  InvoLine.InvoHeadID AS InvoiceNumber,
  InvoLine.Code,
  Item.Description,
  InvoLine.Quantity,
  InvoLine.Price AS Amount
FROM
  Cashier
  INNER JOIN (
    (
      (
        InvoLine
        LEFT JOIN qryXCancelledSales ON (
          InvoLine.CashSysID = qryXCancelledSales.CashSysID
        )
        AND (
          InvoLine.InvoHeadID = qryXCancelledSales.InvoHeadID
        )
      )
      INNER JOIN Item ON InvoLine.Code = Item.ID
    )
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  ) ON Cashier.ID = InvoLine.CashierID
WHERE
  (
    (
      (InvoLine.Quantity)< 0
    )
    AND (
      (qryXCancelledSales.InvoHeadID) Is Null
    )
    AND (
      (InvoLine.TransType)= "S"
    )
    AND (
      (SalesDepartment.Discounts)= False
    )
  )
ORDER BY
  InvoLine.CashierID,
  InvoLine.SetupDate;
