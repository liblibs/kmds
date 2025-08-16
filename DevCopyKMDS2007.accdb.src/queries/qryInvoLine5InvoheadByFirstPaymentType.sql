SELECT
  InvoLine5.CashSysID,
  InvoLine5.InvoHeadID,
  InvoLine5.SetupDate,
  Max(InvoLine5.Code) AS TenderCode
FROM
  InvoLine5
WHERE
  (
    (
      (InvoLine5.TransType)= "T"
    )
  )
GROUP BY
  InvoLine5.CashSysID,
  InvoLine5.InvoHeadID,
  InvoLine5.SetupDate
HAVING
  (
    (
      (
        Max(InvoLine5.Code)
      )< 10
    )
  );
