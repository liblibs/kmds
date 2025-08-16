DELETE InvoLine.SetupDate,
InvoLine.*
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.SetupDate)< #1/1/2016#
    )
  );
