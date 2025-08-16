SELECT
  0 - Sum(
    IIf([Code] = 2, [Price], 0)
  ) AS Checks,
  0 - Sum(
    IIf([Code] = 3, [Price], 0)
  ) AS Cards,
  0 - Sum(
    IIf([Code] = 10, [Price], 0)
  ) AS EFT
FROM
  InvoLine5
WHERE
  (
    (
      (InvoLine5.TransType)= "T"
    )
    AND (
      (InvoLine5.InputCode) Like "Drop*"
      Or (InvoLine5.InputCode)= "Cash"
      Or (InvoLine5.InputCode)= "Fintech"
    )
  )
HAVING
  (
    (
      (
        0 - Sum(
          IIf([Code] = 2, [Price], 0)
        )
      )<> 1
    )
  );
