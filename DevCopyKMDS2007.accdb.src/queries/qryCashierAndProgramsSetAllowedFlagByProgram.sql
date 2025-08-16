PARAMETERS parAllowed Bit,
parProgramName Text (255);
UPDATE
  CashiersAndPrograms
SET
  CashiersAndPrograms.Allowed = [parAllowed]
WHERE
  (
    (
      (
        CashiersAndPrograms.ProgramName
      )= [parProgramName]
    )
  );
