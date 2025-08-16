SELECT
  CashiersAndPrograms.CashierID,
  CashiersAndPrograms.ProgramName,
  CashiersAndPrograms.Allowed
FROM
  CashiersAndPrograms
WHERE
  (
    (
      (CashiersAndPrograms.CashierID)= [parCashierID]
    )
    AND (
      (
        CashiersAndPrograms.ProgramName
      )= [parProgramName]
    )
    AND (
      (CashiersAndPrograms.Allowed)= True
    )
  );
