SELECT
  Cashier.ID,
  Cashier.Name,
  CashiersAndPrograms.Type,
  CashiersAndPrograms.Allowed,
  [Switchboard Items].Argument,
  [Switchboard Items].SwitchboardID,
  [Switchboard Items].ItemNumber,
  [Switchboard Items].ItemText
FROM
  [Switchboard Items]
  INNER JOIN (
    CashiersAndPrograms
    INNER JOIN Cashier ON CashiersAndPrograms.CashierID = Cashier.ID
  ) ON [Switchboard Items].Argument = CashiersAndPrograms.ProgramName
WHERE
  (
    (
      (Cashier.ID)= [Forms]![CashiersAndProgramsSwitchBoard]![cmbCashier]
    )
  )
ORDER BY
  [Switchboard Items].SwitchboardID,
  [Switchboard Items].ItemNumber;
