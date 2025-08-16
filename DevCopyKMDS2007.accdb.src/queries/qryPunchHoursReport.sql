SELECT
  TimeClockPunches.PunchDate,
  TimeClockPunches.CashierID,
  Cashier.Name AS CashierName,
  TimeClockPunches.Punch1,
  TimeClockPunches.Punch2,
  TimeClockPunches.Punch3,
  TimeClockPunches.Punch4,
  (
    DateDiff("n", [Punch1], [Punch2])/ 60
  )+(
    DateDiff("n", [Punch3], [Punch4])/ 60
  ) AS HoursWorked,
  Cashier.PayrollID,
  Cashier.HourlyWage,
  (
    (
      (
        DateDiff("n", [Punch1], [Punch2])/ 60
      )+(
        DateDiff("n", [Punch3], [Punch4])/ 60
      )
    )* [HourlyWage]
  ) AS Wages
FROM
  TimeClockPunches
  INNER JOIN Cashier ON TimeClockPunches.CashierID = Cashier.ID
WHERE
  (
    (
      (TimeClockPunches.PunchDate) Between [Forms]![PrintInventoryReport]![txtDateRangeFrom]
      And [Forms]![PrintInventoryReport]![txtDateRangeTo]
    )
  )
ORDER BY
  TimeClockPunches.PunchDate;
