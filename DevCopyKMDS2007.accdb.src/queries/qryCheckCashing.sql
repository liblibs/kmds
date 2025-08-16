SELECT
  CheckCashing.ID,
  CheckCashing.FirstName,
  CheckCashing.LastName,
  CheckCashing.Address,
  [City] & ", " & [State] & "  " & [zip] AS CityStateZ,
  CheckCashing.BankAccount,
  CheckCashing.Zip
FROM
  CheckCashing
ORDER BY
  CheckCashing.FirstName;
