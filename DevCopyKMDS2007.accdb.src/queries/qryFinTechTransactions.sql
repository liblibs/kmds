SELECT
  InvoLine5.InvoHeadID,
  InvoLine5.ID,
  InvoLine5.CashSysID,
  InvoLine5.SetupDate,
  InvoLine5.Code,
  InvoLine5.TransType,
  InvoLine5.CashierID
FROM
  InvoLine5
WHERE
  (
    (
      (InvoLine5.Code)= 10
    )
    AND (
      (InvoLine5.TransType)= "T"
    )
  );
