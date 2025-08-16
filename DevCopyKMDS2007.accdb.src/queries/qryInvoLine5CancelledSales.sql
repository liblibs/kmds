SELECT
  InvoLine5.CashierID,
  InvoLine5.SetupDate,
  InvoLine5.InvoHeadID,
  InvoLine5.CashSysID,
  InvoLine5.ID,
  InvoLine5.Id\2 AS MaxID,
  InvoLine5.Code,
  InvoLine5.InputCode,
  Sysfile.CancelTenderID
FROM
  InvoLine5,
  Sysfile
WHERE
  (
    (
      (InvoLine5.Code)= [CancelTenderID]
    )
  )
ORDER BY
  InvoLine5.CashierID,
  InvoLine5.SetupDate;
