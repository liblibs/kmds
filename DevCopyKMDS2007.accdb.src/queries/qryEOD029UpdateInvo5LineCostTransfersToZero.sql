UPDATE
  InvoLine5
SET
  InvoLine5.Cost = 0
WHERE
  (
    (
      (InvoLine5.TransType)= "X"
    )
  );
