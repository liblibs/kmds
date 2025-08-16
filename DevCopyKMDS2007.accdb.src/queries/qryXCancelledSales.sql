SELECT
  InvoLine.InvoHeadID,
  InvoLine.CashSysID
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.InputCode)= "Cancelled Sales"
    )
  )
GROUP BY
  InvoLine.InvoHeadID,
  InvoLine.CashSysID;
