SELECT
  TimeClockPunches.CashierID,
  TimeClockPunches.PunchDate,
  TimeClockPunches.Punch1,
  TimeClockPunches.Punch2,
  TimeClockPunches.Punch3,
  TimeClockPunches.Punch4,
  TimeClockPunches.SetupDate,
  TimeClockPunches.ChangeDate,
  (
    DateDiff("n", [Punch1], [Punch2])/ 60
  )+(
    DateDiff("n", [Punch3], [Punch4])/ 60
  ) AS HoursWorked,
  (
    (
      (
        DateDiff("n", [Punch1], [Punch2])/ 60
      )+(
        DateDiff("n", [Punch3], [Punch4])/ 60
      )
    )* Cashier.HourlyWage
  ) AS Wages
FROM
  TimeClockPunches
  INNER JOIN Cashier ON TimeClockPunches.CashierID = Cashier.ID;
