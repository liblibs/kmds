SELECT
  Cashier.ID,
  Cashier.Name,
  CashiersAndPrograms.Allowed,
  CashiersAndPrograms.Type,
  CashiersAndPrograms.ProgramName
FROM
  CashiersAndPrograms
  INNER JOIN Cashier ON CashiersAndPrograms.CashierID = Cashier.ID
WHERE
  (
    (
      (
        CashiersAndPrograms.ProgramName
      )= [Forms]![CashierRights]![LabelMenuOption].[Tag]
    )
  )
ORDER BY
  Cashier.ID;
