UPDATE
  InvoLine
SET
  InvoLine.SalesGroup = Left(
    [InputCode],
    InStr([InputCode], "-")-1
  )
WHERE
  (
    (
      (InvoLine.SalesGroup)<> Left(
        [InputCode],
        InStr([InputCode], "-")-1
      )
    )
    AND (
      (InvoLine.TransType)= "P"
    )
  );
