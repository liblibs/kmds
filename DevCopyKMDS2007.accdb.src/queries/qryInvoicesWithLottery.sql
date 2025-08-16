SELECT
  InvoLine.InvoHeadID,
  InvoLine.CashSysID,
  InvoLine.SetupDate
FROM
  InvoLine
  INNER JOIN Item ON InvoLine.Code = Item.ID
WHERE
  (
    (
      (Item.SalesDepartmentID)= 7
      Or (Item.SalesDepartmentID)= 11
    )
  )
GROUP BY
  InvoLine.InvoHeadID,
  InvoLine.CashSysID,
  InvoLine.SetupDate;
