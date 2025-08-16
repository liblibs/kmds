DELETE InvoLine.Quantity,
InvoLine.TransType
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.Quantity)= 0
    )
    AND (
      (InvoLine.TransType)<> "R"
    )
  );
