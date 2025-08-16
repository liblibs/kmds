DELETE InvoLine.TransType,
*
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.TransType)= "Z"
    )
  );
