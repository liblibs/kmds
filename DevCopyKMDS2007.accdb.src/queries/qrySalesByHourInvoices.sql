SELECT
  Format([SetupDate], "mm/dd/yyyy") AS BusinessDate,
  InvoLine.CashSysID,
  InvoLine.InvoHeadID
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.TransType)= "S"
    )
  )
GROUP BY
  Format([SetupDate], "mm/dd/yyyy"),
  InvoLine.CashSysID,
  InvoLine.InvoHeadID;
