PARAMETERS parCashierID Short,
parProgramName Text (255);
SELECT
  Cashier.ID AS CashierID,
  Programs.Name AS ProgramName,
  "Program" AS Type
FROM
  Cashier,
  Programs
WHERE
  (
    (
      (Cashier.ID)= [parCashierID]
    )
    AND (
      (Programs.Name)= [parProgramName]
    )
  );
