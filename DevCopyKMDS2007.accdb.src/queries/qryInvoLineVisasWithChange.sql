SELECT
  InvoLine.SetupDate,
  InvoLine.CashSysID,
  InvoLine.InvoHeadID,
  qryInvoLineVisas.Price,
  InvoLine.Price,
  InvoLine.InputCode,
  InvoLine.CashierID
FROM
  qryInvoLineVisas
  INNER JOIN InvoLine ON (
    qryInvoLineVisas.InvoHeadID = InvoLine.InvoHeadID
  )
  AND (
    qryInvoLineVisas.CashSysID = InvoLine.CashSysID
  )
WHERE
  (
    (
      (InvoLine.Price)<> 0
    )
    AND (
      (InvoLine.InputCode)= "Change"
    )
    AND (
      (InvoLine.TransType)= "T"
    )
    AND (
      (InvoLine.Code)<> 3
    )
  )
ORDER BY
  InvoLine.SetupDate,
  InvoLine.CashSysID,
  InvoLine.InvoHeadID;
