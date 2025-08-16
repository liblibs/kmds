TRANSFORM First(CashiersAndPrograms.Allowed) AS FirstOfAllowed
SELECT
  CashiersAndPrograms.ProgramName,
  First(CashiersAndPrograms.Allowed) AS [Total Of Allowed]
FROM
  CashiersAndPrograms
GROUP BY
  CashiersAndPrograms.ProgramName PIVOT CashiersAndPrograms.CashierID;
