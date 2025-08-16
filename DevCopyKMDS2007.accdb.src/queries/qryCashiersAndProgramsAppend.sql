INSERT INTO CashiersAndPrograms (CashierID, ProgramName, Type)
SELECT
  qryCashiersAndPrograms.CashierID,
  qryCashiersAndPrograms.ProgramName,
  qryCashiersAndPrograms.Type
FROM
  qryCashiersAndPrograms
  LEFT JOIN CashiersAndPrograms ON (
    qryCashiersAndPrograms.ProgramName = CashiersAndPrograms.ProgramName
  )
  AND (
    qryCashiersAndPrograms.CashierID = CashiersAndPrograms.CashierID
  )
WHERE
  (
    (
      (CashiersAndPrograms.CashierID) Is Null
    )
  );
