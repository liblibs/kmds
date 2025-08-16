SELECT
  InvoLine5.InvoHeadID,
  InvoLine5.CashSysID,
  InvoLine5.SetupDate,
  First(InvoLine5.Code) AS TenderCode
FROM
  InvoLine5
WHERE
  (
    (
      (InvoLine5.TransType)= "T"
    )
  )
GROUP BY
  InvoLine5.InvoHeadID,
  InvoLine5.CashSysID,
  InvoLine5.SetupDate
ORDER BY
  First(InvoLine5.Code);
