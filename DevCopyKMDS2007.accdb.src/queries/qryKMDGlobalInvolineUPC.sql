SELECT
  InvoLine.Code,
  InvoLine.InputCode,
  Last(InvoLine.SetupDate) AS LastOfSetupDate,
  Sysfile.AgentNumber
FROM
  Sysfile,
  InvoLine
  INNER JOIN UPCItem ON InvoLine.InputCode = UPCItem.ID
WHERE
  (
    (
      (
        Len([InputCode])
      )> 5
    )
  )
GROUP BY
  InvoLine.Code,
  InvoLine.InputCode,
  Sysfile.AgentNumber
HAVING
  (
    (
      (InvoLine.Code) Between 1000
      And 99999
    )
  );
