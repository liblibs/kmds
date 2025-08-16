SELECT
  InvoLine.InputCode,
  Max(InvoLine.SetupDate) AS MaxOfSetupDate INTO InvoLineUPCMaxDate
FROM
  InvoLine
GROUP BY
  InvoLine.InputCode,
  InvoLine.TransType,
  Len([INPUTCODE])
HAVING
  (
    (
      (InvoLine.TransType)= "S"
    )
    AND (
      (
        Len([INPUTCODE])
      )> 7
    )
  );
