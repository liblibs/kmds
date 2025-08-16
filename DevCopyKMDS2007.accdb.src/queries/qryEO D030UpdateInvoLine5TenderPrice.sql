UPDATE
  InvoLine5
SET
  InvoLine5.Price = 0 - [Price],
  InvoLine5.Quantity = 0 - [Quantity]
WHERE
  (
    (
      (InvoLine5.TransType)= "T"
    )
  );
