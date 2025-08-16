UPDATE
  InvoLine,
  qryDaysToRetain
SET
  InvoLine.TransType = "Z"
WHERE
  (
    (
      (InvoLine.SetupDate)< Now()- [qrydaystoRetain].[days]
    )
  );
