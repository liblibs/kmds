SELECT
  Cashier.ID AS CashierID,
  Programs.Name AS ProgramName,
  "Program" AS Type
FROM
  Cashier,
  Programs;
