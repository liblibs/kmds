SELECT
  TimeClockPunches.CashierID,
  Cashier.Name
FROM
  TimeClockPunches
  LEFT JOIN Cashier ON TimeClockPunches.CashierID = Cashier.ID
GROUP BY
  TimeClockPunches.CashierID,
  Cashier.Name;
