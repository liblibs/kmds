SELECT
  InvoLine.InvoHeadID,
  InvoLine.CashSysID,
  InvoLine.TransType,
  InvoLine.Code,
  InvoLine.Price,
  InvoLine.InputCode
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.TransType)= "T"
    )
    AND (
      (InvoLine.Code)= 3
    )
    AND (
      (InvoLine.Price)<> 0
    )
  );
