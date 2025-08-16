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
      (Item.ID)= 997006
    )
  )
GROUP BY
  InvoLine.InvoHeadID,
  InvoLine.CashSysID,
  InvoLine.SetupDate;
