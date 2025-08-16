UPDATE
  InvoLine4
SET
  InvoLine4.TransType = "S"
WHERE
  (
    (
      ([InvoLine4].[TransType]) Is Null
    )
  );
