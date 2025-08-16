SELECT
  InvoLine5.InvoHeadID,
  InvoLine5.CashSysID,
  InvoLine5.TransType,
  InvoLine5.InputCode AS InvoiceNumber
FROM
  InvoLine5
GROUP BY
  InvoLine5.InvoHeadID,
  InvoLine5.CashSysID,
  InvoLine5.TransType,
  InvoLine5.InputCode
HAVING
  (
    (
      (InvoLine5.TransType)= "D"
    )
  );
